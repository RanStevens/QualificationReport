# QualificationReport

A Go project for generating qualification reports and working with Snowflake data.

## Overview

This repository contains a Go application that reads environment configuration, connects to Snowflake, and processes report data into Excel format.

## Requirements

- Go 1.26+
- A Snowflake account
- Environment variables configured in `.env`

## Setup

1. Copy the example environment file:
   ```bash
   cp .env.example .env
   ```
2. Update `.env` with your Snowflake credentials and any required configuration.
3. Install dependencies:
   ```bash
   go mod download
   ```

## Run

```bash
go run Main.go
```

## Development

Format code:

```bash
go fmt ./...
```

Run tests:

```bash
go test ./...
```

## CI

GitHub Actions is configured to run `go test ./...` on push and pull requests.
