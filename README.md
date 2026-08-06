# SAP Purchase Requisition Automation

> Internal SAP GUI automation developed to streamline Purchase Requisition (PR) validation, creation, and removal for the Purchasing team.

---

# Overview

This project was developed to automate repetitive Purchase Requisition (PR) activities performed by buyers in SAP.

Before this automation, users manually validated Plan Orders, checked business rules, and created or removed Purchase Requisitions through multiple SAP transactions. These repetitive tasks were time-consuming and susceptible to human error.

The solution uses Excel VBA together with SAP GUI Scripting to automate the workflow while maintaining existing business validation rules.

> **Status:** Production

---

# Business Problem

The Purchasing team frequently needed to:

- Validate Plan Orders before creating Purchase Requisitions.
- Check supplier-specific business rules.
- Create Purchase Requisitions manually.
- Remove unnecessary Purchase Requisitions.

The manual process required repetitive SAP navigation and multiple validation steps, resulting in unnecessary effort and inconsistent execution.

---

# Solution

This automation streamlines the Purchase Requisition workflow by:

- Validating Plan Orders before PR creation.
- Applying business validation rules.
- Automating Purchase Requisition creation.
- Automating Purchase Requisition removal.
- Reducing repetitive SAP interactions.

The objective is to standardize the workflow while improving efficiency for buyers.

---

# Features

- Purchase Requisition validation
- Purchase Requisition creation
- Purchase Requisition removal
- SAP GUI automation
- Business rule validation
- Excel-based user interface

---

# Technologies Used

| Category | Technologies |
|----------|--------------|
| Language | Excel VBA |
| ERP | SAP MM |
| Automation | SAP GUI Scripting |
| Office | Microsoft Excel |

---

# Workflow

```text
User Input (Excel)
        │
        ▼
Validate Plan Orders
        │
        ▼
Apply Business Rules
        │
        ▼
Create / Remove Purchase Requisition
        │
        ▼
Return Processing Result
```

---

# Business Impact

This automation helps the Purchasing team by:

- Reducing repetitive SAP processing.
- Standardizing Purchase Requisition validation.
- Improving consistency during PR processing.
- Supporting buyers with a repeatable workflow.

---

# My Role

I was responsible for:

- Understanding business requirements.
- Designing the automation workflow.
- Developing the VBA automation.
- Integrating with SAP GUI Scripting.
- Maintaining and enhancing the solution after deployment.

---

# Lessons Learned

Developing this project reinforced several important principles:

- Business process knowledge is as important as programming.
- Validation should occur before automation execution.
- Repetitive SAP tasks can be standardized through automation.
- Modular design improves long-term maintainability.

---

# Future Improvements

If rebuilding this project today, I would:

- Rebuild the solution using Python.
- Replace Excel-based validation with pandas.
- Separate SAP operations into reusable modules.
- Introduce centralized logging.
- Support parallel SAP sessions.
- Store configuration in external configuration files.
- Improve maintainability through a modular architecture.

---

# Repository Status

**Current Status:** Production

This repository represents the first production version of the automation. Future enhancements will focus on improving scalability, maintainability, and reusability using Python.
