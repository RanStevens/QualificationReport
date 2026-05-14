package main

import (
	"database/sql"
	"encoding/base64"
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/joho/godotenv"
	sf "github.com/snowflakedb/gosnowflake"
	"github.com/xuri/excelize/v2"
)

// excelMoveFirstColsToEnd is used only when excelOutputOrder is nil: that many leading
// result columns are written after all other columns in Excel (same as before).
const excelMoveFirstColsToEnd = 2

// excelOutputOrder defines the Excel column layout. If non-nil, len must match the
// query result each run, and excelOutputOrder[i] is the source column index (0-based,
// same order as rows.Columns / rows.Scan) placed at Excel column i.
// Example (5 columns): []int{4, 1, 2, 3, 0} puts original col 4 first, then 1,2,3, then col 0 last.
// Leave nil to use excelMoveFirstColsToEnd instead.
// var excelOutputOrder []int
var excelOutputOrder = []int{2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 0, 1}

// buildRotateFirstNPermutation returns perm where Excel column i shows source column perm[i]:
// sources k..n-1 first, then sources 0..k-1 (move first k columns to the end).
func buildRotateFirstNPermutation(n, k int) []int {
	p := make([]int, n)
	if k <= 0 || n == 0 || k >= n {
		for i := range p {
			p[i] = i
		}
		return p
	}
	for i := 0; i < n-k; i++ {
		p[i] = i + k
	}
	for i := 0; i < k; i++ {
		p[n-k+i] = i
	}
	return p
}

func validateExcelPermutation(perm []int, n int) ([]int, error) {
	if len(perm) != n {
		return nil, fmt.Errorf("excelOutputOrder length %d != result column count %d", len(perm), n)
	}
	seen := make([]bool, n)
	for i, src := range perm {
		if src < 0 || src >= n {
			return nil, fmt.Errorf("excelOutputOrder[%d]=%d out of range for n=%d", i, src, n)
		}
		if seen[src] {
			return nil, fmt.Errorf("excelOutputOrder: duplicate source column index %d", src)
		}
		seen[src] = true
	}
	return perm, nil
}

func excelPermutation(n int) ([]int, error) {
	if excelOutputOrder != nil {
		return validateExcelPermutation(excelOutputOrder, n)
	}
	return buildRotateFirstNPermutation(n, excelMoveFirstColsToEnd), nil
}

// reorderByPermutation builds out[i] = orig[perm[i]].
func reorderByPermutation(orig []interface{}, perm []int) []interface{} {
	out := make([]interface{}, len(perm))
	for i, src := range perm {
		out[i] = orig[src]
	}
	return out
}

// Config holds Snowflake connection settings and report output options.
type Config struct {
	User       string
	Password   string
	Account    string
	Database   string
	Schema     string
	Warehouse  string
	Role       string
	OutputPath string
}

// readSnowflakePassword returns the raw password for Snowflake.
// Precedence: SNOWFLAKE_PASSWORD_FILE (file contents, trimmed), else
// SNOWFLAKE_PASSWORD_B64 (standard Base64 of UTF-8 password), else SNOWFLAKE_PASSWORD.
// Do not put URL-encoded passwords in .env — supply the same raw string Snowflake expects.
func readSnowflakePassword() (string, error) {
	if path := strings.TrimSpace(os.Getenv("SNOWFLAKE_PASSWORD_FILE")); path != "" {
		b, err := os.ReadFile(path)
		if err != nil {
			return "", fmt.Errorf("read SNOWFLAKE_PASSWORD_FILE %q: %w", path, err)
		}
		return strings.TrimSpace(string(b)), nil
	}
	if b64 := strings.TrimSpace(os.Getenv("SNOWFLAKE_PASSWORD_B64")); b64 != "" {
		raw, err := base64.StdEncoding.DecodeString(b64)
		if err != nil {
			return "", fmt.Errorf("decode SNOWFLAKE_PASSWORD_B64: %w", err)
		}
		return string(raw), nil
	}
	return strings.TrimSpace(os.Getenv("SNOWFLAKE_PASSWORD")), nil
}

// LoadConfig reads configuration from the process environment.
// Required: SNOWFLAKE_USER; password from SNOWFLAKE_PASSWORD and/or SNOWFLAKE_PASSWORD_FILE
// and/or SNOWFLAKE_PASSWORD_B64; SNOWFLAKE_ACCOUNT, SNOWFLAKE_DATABASE, SNOWFLAKE_SCHEMA,
// SNOWFLAKE_WAREHOUSE.
// Optional: SNOWFLAKE_ROLE, REPORT_OUTPUT_PATH (default QualificationReport.xlsx).
func LoadConfig() (Config, error) {
	var missing []string
	req := func(name string) string {
		v := strings.TrimSpace(os.Getenv(name))
		if v == "" {
			missing = append(missing, name)
		}
		return v
	}

	pw, err := readSnowflakePassword()
	if err != nil {
		return Config{}, err
	}
	if pw == "" {
		missing = append(missing, "SNOWFLAKE_PASSWORD (or SNOWFLAKE_PASSWORD_FILE / SNOWFLAKE_PASSWORD_B64)")
	}

	cfg := Config{
		User:      req("SNOWFLAKE_USER"),
		Password:  pw,
		Account:   req("SNOWFLAKE_ACCOUNT"),
		Database:  req("SNOWFLAKE_DATABASE"),
		Schema:    req("SNOWFLAKE_SCHEMA"),
		Warehouse: req("SNOWFLAKE_WAREHOUSE"),
		Role:      strings.TrimSpace(os.Getenv("SNOWFLAKE_ROLE")),
	}
	if len(missing) > 0 {
		return Config{}, fmt.Errorf("missing required environment variables: %s (see .env.example)", strings.Join(missing, ", "))
	}

	out := strings.TrimSpace(os.Getenv("REPORT_OUTPUT_PATH"))
	if out == "" {
		out = "QualificationReport.xlsx"
	}
	cfg.OutputPath = out

	return cfg, nil
}

// SnowflakeDSN builds the driver DSN the same way gosnowflake does internally (user/password
// escaping, account+region split, query params). Prefer this over hand-rolled fmt.Sprintf DSNs.
func (c Config) SnowflakeDSN() (string, error) {
	sfc := &sf.Config{
		User:      c.User,
		Password:  c.Password,
		Account:   c.Account,
		Database:  c.Database,
		Schema:    c.Schema,
		Warehouse: c.Warehouse,
		Role:      c.Role,
	}
	return sf.DSN(sfc)
}

// loadLocalEnv loads key=value pairs into the process environment.
// Uses DOTENV_PATH if set; otherwise looks for ".env" in the current working directory.
func loadLocalEnv() {
	path := strings.TrimSpace(os.Getenv("DOTENV_PATH"))
	if path == "" {
		path = ".env"
	}
	info, err := os.Stat(path)
	if errors.Is(err, os.ErrNotExist) {
		if strings.TrimSpace(os.Getenv("DOTENV_PATH")) != "" {
			log.Fatalf("DOTENV_PATH is set to %q but that file does not exist", path)
		}
		return
	}
	if err != nil {
		log.Fatalf("stat env file %q: %v", path, err)
	}
	if info.IsDir() {
		log.Fatalf("env file path %q is a directory, not a file", path)
	}
	// Overload (not Load): standard godotenv.Load skips keys already present in the
	// environment. IDEs often inject empty SNOWFLAKE_* placeholders, which would block .env.
	if err := godotenv.Overload(path); err != nil {
		abs, _ := filepath.Abs(path)
		log.Fatalf("loading env file %s: %v", abs, err)
	}
}

func explainMissingEnv(err error) {
	wd, _ := os.Getwd()
	log.Printf("working directory: %s", wd)
	path := strings.TrimSpace(os.Getenv("DOTENV_PATH"))
	if path == "" {
		path = ".env"
	}
	abs, _ := filepath.Abs(path)
	if _, e := os.Stat(path); errors.Is(e, os.ErrNotExist) {
		log.Printf("no file at %s — either cd into the project directory that contains .env, or set DOTENV_PATH to the full path of your .env file", abs)
	} else if e != nil {
		log.Printf("could not stat %s: %v", abs, e)
	} else {
		log.Printf("loaded %s (values from this file override same-named process env vars); if variables are still missing, check for typos in names (must match exactly) and non-empty values after the = sign", abs)
	}
	log.Fatal(err)
}

func main() {
	loadLocalEnv()

	cfg, err := LoadConfig()
	if err != nil {
		explainMissingEnv(err)
	}
	dsn, err := cfg.SnowflakeDSN()
	if err != nil {
		log.Fatal("snowflake DSN:", err)
	}

	db, xerr := sql.Open("snowflake", dsn)
	if xerr != nil {
		log.Fatal("Error creating connection:", xerr)
	}
	defer db.Close()
	// Verify the connection is actually alive
	xerr = db.Ping()
	if xerr != nil {
		log.Fatal("Could not establish connection:", xerr)
	}

	fmt.Println("Successfully connected to Snowflake!")

	var version string
	xerr = db.QueryRow("SELECT CURRENT_VERSION()").Scan(&version)
	if xerr != nil {
		log.Fatal(xerr)
	}
	fmt.Printf("Snowflake Version: %s\n", version)

	rows, err := db.Query("CALL CCR_REPT_MCGEN()")
	if err != nil {
		log.Fatal(err)
	}
	defer rows.Close()

	columns, err := rows.Columns()
	if err != nil {
		log.Fatal("columns:", err)
	}

	perm, err := excelPermutation(len(columns))
	if err != nil {
		log.Fatal("excel column order:", err)
	}

	outPath := cfg.OutputPath
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	sheet := f.GetSheetName(0)

	header := make([]interface{}, len(columns))
	for i, c := range columns {
		header[i] = c
	}
	header = reorderByPermutation(header, perm)
	if err := f.SetSheetRow(sheet, "A1", &header); err != nil {
		log.Fatal("write header:", err)
	}

	rowNum := 2
	for rows.Next() {
		var Run_Unique_ID_1 string
		var Run_Unique_ID_2 string
		var LvlStr string
		var Pin int32
		var Street string
		var FullName string
		var Bonus_Diff float64
		var Ttl_Bonus float64
		var Basic_Bonus_Diff float64
		var Inf_Bonus_Diff float64
		var Dia_Bonus_Diff float64
		var LM_Rank int
		var Rank int
		var AutoQual int
		var PV float32
		var GV_Diff float32
		var GV float32
		var Basic_Legs int
		var Bronze_Legs int
		var Dia_Legs int
		var Basic_Bonus float64
		var Inf_Bonus float64
		var Dia_Bonus float64
		var Basic_Legs_Diff int
		var Bronze_Legs_Diff int
		var Dia_Legs_Diff int
		var PWV int
		var Qual_Pay_Lvl int
		var Period_Status int

		err = rows.Scan(&Run_Unique_ID_1, &Run_Unique_ID_2, &LvlStr, &Pin, &Street, &FullName, &Bonus_Diff, &Ttl_Bonus, &Basic_Bonus_Diff, &Inf_Bonus_Diff, &Dia_Bonus_Diff, &LM_Rank, &Rank, &AutoQual, &PV, &GV_Diff, &GV, &Basic_Legs, &Bronze_Legs, &Dia_Legs, &Basic_Bonus, &Inf_Bonus, &Dia_Bonus, &Basic_Legs_Diff, &Bronze_Legs_Diff, &Dia_Legs_Diff, &PWV, &Qual_Pay_Lvl, &Period_Status)
		if err != nil {
			log.Fatal(err)
		}

		values := []interface{}{
			Run_Unique_ID_1, Run_Unique_ID_2, LvlStr, Pin, Street, FullName,
			Bonus_Diff, Ttl_Bonus, Basic_Bonus_Diff, Inf_Bonus_Diff, Dia_Bonus_Diff,
			LM_Rank, Rank, AutoQual, PV, GV_Diff, GV,
			Basic_Legs, Bronze_Legs, Dia_Legs, Basic_Bonus, Inf_Bonus, Dia_Bonus,
			Basic_Legs_Diff, Bronze_Legs_Diff, Dia_Legs_Diff, PWV, Qual_Pay_Lvl, Period_Status,
		}
		values = reorderByPermutation(values, perm)
		if len(values) != len(columns) {
			log.Fatalf("column count mismatch: scan has %d values, result set has %d columns", len(values), len(columns))
		}
		startCell, err := excelize.CoordinatesToCellName(1, rowNum)
		if err != nil {
			log.Fatal(err)
		}
		if err := f.SetSheetRow(sheet, startCell, &values); err != nil {
			log.Fatal("write row:", err)
		}
		rowNum++
	}
	if err := rows.Err(); err != nil {
		log.Fatal("rows:", err)
	}
	if err := f.SaveAs(outPath); err != nil {
		log.Fatal("save:", err)
	}
	fmt.Printf("Wrote %d data rows to %s\n", rowNum-2, outPath)
}
