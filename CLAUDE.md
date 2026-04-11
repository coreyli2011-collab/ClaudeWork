# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## About the User

The user is a manufacturing and NVH (Noise, Vibration & Harshness) engineer in the automotive industry with no programming or Python background. Do not assume any coding knowledge.

**Always:**
- Explain what you are doing and why in plain, non-technical language before writing or running anything
- Ask for explicit approval before executing any script, command, or file operation
- Relate technical concepts to engineering or everyday analogies where helpful

## Purpose

This workspace is for automating repetitive tasks in an automotive engineering context, including:
- File and data organisation, cleanup, and transformation
- Analysis and report generation
- Workflow automation between apps such as Siemens Testlab and Microsoft Office (Excel, Word, PowerPoint)

## Directory Layout

- `data-cleanup/` — scripts and utilities for cleaning or transforming data
- `file-automation/` — scripts for automating file operations
- `reports/` — generated output and report files
- `scripts/` — general-purpose utility scripts

## Project Folder Naming Convention

Major project folders use the prefix `P-` so they group together and are easy to distinguish from utility folders.

- `P-TestLab/` — Siemens Testlab automation projects
- `P-PowerPoint/` — PowerPoint report generation projects
- `P-WebApps/` — web apps and interactive tools

**Rules:**
- Only create a `P-` folder when a new project actually starts — do not create empty folders in advance
- Utility folders (`data-cleanup/`, `file-automation/`, `scripts/`, `reports/`) keep their existing names
- `david-adventure` stays in the root folder since its link is already shared publicly
