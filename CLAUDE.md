# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an **ABAP (SAP)** application for bulk updating business partner email addresses and address notes in SAP ECC/NetWeaver systems. It supports two modes: batch updates from an uploaded Excel file, and single business partner lookup/update via RFC from a remote SAP system.

## Running the Program

This code runs inside a SAP system — there are no local build, lint, or test commands.

- **Execute:** Via transaction `SE38` (ABAP Editor) or `SE80` (Object Navigator) using program name `ZMM_UPDATE_EMAIL_ADDR`
- **Test mode:** The selection screen has a test mode flag that validates changes without committing to the database

## Architecture

The application uses **MVP (Model-View-Presenter)** pattern across four files in `update_bp_email/`:

- **`ZMM_UPDATE_EMAIL_ADDR.abap`** — Report entry point; defines the selection screen and invokes the presenter
- **`ZMM_CL_BP_UPD_PRESENTER.abap`** — Orchestrates the two execution paths: `run()` for batch Excel processing and `run_single_bp()` for RFC-based single BP lookup
- **`ZMM_CL_BP_UPD_MODEL.abap`** — All business logic: Excel parsing (`ALSM_EXCEL_TO_INTERNAL_TABLE`), BAPI calls for updates (`BAPI_BUPA_ADDRESS_CHANGE`, `BAPI_TRANSACTION_COMMIT/ROLLBACK`), and RFC calls (`RFC_READ_TABLE`) to query remote SAP tables (LFA1, KNA1, ADR6, ADRT)
- **`ZMM_CL_BP_UPD_VIEW.abap`** — UI layer: file dialog, ALV results grid display, popup messages, and column configuration

## Key SAP Integration Points

| Component | Purpose |
|-----------|---------|
| `ALSM_EXCEL_TO_INTERNAL_TABLE` | Reads uploaded Excel files |
| `BAPI_BUPA_ADDRESS_CHANGE` | Updates BP address/email |
| `RFC_READ_TABLE` | Generic remote table reader for cross-system lookups |
| LFA1 / KNA1 | Vendor / Customer master tables |
| ADR6 / ADRT | Email addresses / Address notes |

## Data Structures

- `ty_itab` — Main processing table: business partner number, email, notes, update flags
- `tt_errorlog` — Result/error log used to populate the ALV output grid
