# Camunda RPA Test Automation

This repository contains a Camunda-based RPA test automation setup for validating desktop and web workflows.

The project combines:

- Camunda 8 Self-Managed
- Camunda RPA Worker
- Robot Framework
- Python / pywinauto
- Playwright
- Spring Boot
- CSV-based test result persistence

## Architecture

The main Camunda environment runs on the primary PC.

RPA workers can run on multiple Windows machines, for example:

- Windows VM
- Laptop

Each worker connects to the same Camunda instance and can be selected through worker labels.

Example worker labels:

```text
vm
laptop
outlook
default
