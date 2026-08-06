# SAP Purchase Requisition Automation

## Overview

This project automates the validation, creation, and removal of Purchase Requisitions (PR) in SAP.

It was developed to reduce repetitive manual work performed by buyers while improving consistency during PR processing.

The solution combines Excel VBA with SAP GUI Scripting to automate business workflows based on predefined validation rules.

---

## Business Problem

Before this automation was introduced, buyers manually:

- Validated Plan Orders
- Verified business rules
- Created Purchase Requisitions
- Removed unnecessary Purchase Requisitions

The process involved repetitive SAP navigation and manual validation, making it time-consuming and increasing the risk of inconsistent processing.

---

## Solution

This automation streamlines the Purchase Requisition process by:

- Validating Plan Orders before PR creation
- Applying business validation rules
- Automating PR creation
- Automating PR removal

---

## Technologies Used

- Excel VBA
- SAP GUI Scripting
- Microsoft Excel
- SAP MM

---

## Business Impact

- Reduced repetitive manual SAP processing
- Standardized Purchase Requisition validation
- Improved consistency during PR creation

---

## Future Improvements

If rebuilding this project today, I would:

- Rebuild the solution using Python.
- Replace Excel-based validation with pandas.
- Separate SAP operations into reusable modules.
- Introduce centralized logging.
- Support parallel SAP sessions.
