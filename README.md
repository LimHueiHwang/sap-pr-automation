# SAP Purchase Requisition Automation

![Production](https://img.shields.io/badge/Status-Production-success)
![SAP MM](https://img.shields.io/badge/SAP-MM-blue)
![Excel VBA](https://img.shields.io/badge/VBA-Excel-green)

> Production SAP GUI automation developed to streamline Purchase Requisition (PR) creation for the Purchasing team, with pre-creation Plan Order and Sales Agreement checks and a separate PR closure workflow.

---

## 📌 Project Summary

This project automates Purchase Requisition (PR) creation for the Purchasing team using Microsoft Excel VBA and SAP GUI Scripting.

The main workflow checks the available Plan Order quantity and determines whether the part is a Sales Agreement (SA) part before creating the Purchase Requisition.

For SA parts, an additional manual check is required to confirm that sufficient Plan Order quantity is available before the PR is created.

The tool also includes a separate workflow for closing or fixing existing Purchase Requisitions.

**Project Status:** Production

---

## 🎯 Business Problem

Before this automation was introduced, buyers manually performed several checks and SAP activities before creating Purchase Requisitions.

The process required users to:

- Check available Plan Order quantity.
- Determine whether the part is a Sales Agreement (SA) part.
- Perform an additional Plan Order quantity check for SA parts.
- Create the Purchase Requisition in SAP.
- Separately process existing Purchase Requisitions that needed to be closed or fixed.

These activities involved repetitive SAP navigation, manual checking, and data entry.

---

## 💡 Solution

The automation streamlines the Purchase Requisition process by:

- Checking Plan Order quantity before PR creation.
- Identifying whether the part is a Sales Agreement (SA) part.
- Automatically proceeding with PR creation for non-SA parts.
- Requiring a manual Plan Order sufficiency check for SA parts.
- Stopping the PR creation process when the required condition is not met.
- Creating the Purchase Requisition through SAP GUI Scripting.
- Returning SAP processing messages and PR information to Excel.
- Providing a separate workflow for closing or fixing existing Purchase Requisitions.

The automation was developed around the existing Purchasing workflow and SAP process.

---

## ✨ Core Features

### 1. Purchase Requisition Creation

The primary function of the automation is to create Purchase Requisitions in SAP.

The workflow:

1. Accepts the required information through Excel.
2. Checks the available Plan Order quantity.
3. Determines whether the part is a Sales Agreement (SA) part.
4. Proceeds directly to PR creation for non-SA parts.
5. Requires an additional manual Plan Order quantity check for SA parts.
6. Creates the Purchase Requisition when the required condition is satisfied.
7. Returns the SAP processing result to Excel.

---

### 2. Pre-Creation Validation

Before creating the PR, the automation performs the following checks:

- Plan Order quantity.
- Sales Agreement (SA) part condition.

The processing path differs depending on whether the part is an SA part.

#### Non-SA Part

```text
Check Plan Order
       ↓
   Not SA Part
       ↓
   Create PR
