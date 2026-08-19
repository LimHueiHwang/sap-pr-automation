# SAP Purchase Requisition Automation

![Production](https://img.shields.io/badge/Status-Production%20%2F%20SAP%20Upgrade%20Review-orange)
![SAP MM](https://img.shields.io/badge/SAP-MM-blue)
![Excel VBA](https://img.shields.io/badge/VBA-Excel-green)
![SAP GUI Scripting](https://img.shields.io/badge/SAP%20GUI-Scripting-orange)

> Excel VBA and SAP GUI automation developed to support Purchase Requisition creation, Plan Order / Sales Agreement validation, and Purchase Requisition Close/Fix activities for the Purchasing team.

---

## Overview

This project is a production Excel VBA automation developed to streamline repetitive Purchase Requisition (PR) activities performed in SAP.

The solution combines:

- Microsoft Excel
- Excel VBA
- SAP GUI Scripting
- SAP Purchasing business rules

The automation supports three functional areas:

1. **Plan Order / Sales Agreement Validation**
2. **Purchase Requisition Creation**
3. **Purchase Requisition Close/Fix**

### Primary Purpose

The **primary purpose of this automation is Purchase Requisition creation**.

Plan Order and Sales Agreement (SA) validation supports the PR creation process by helping determine whether the required conditions are met before proceeding.

The **Close/Fix PR process is a separate workflow** used to update existing Purchase Requisitions.

**Project Status:** Production / SAP Upgrade Compatibility Review

---

# Business Problem

Purchase Requisition processing involved repetitive activities between Excel and SAP.

Purchasing users needed to:

- Review material and quantity requirements.
- Check available Plan Order quantity.
- Determine whether a material is a Sales Agreement (SA) part.
- Perform additional checking when an SA part is identified.
- Navigate through SAP to process the required information.
- Create Purchase Requisitions.
- Review SAP processing messages.
- Process existing Purchase Requisitions that required Close/Fix actions.

Performing these activities manually required repeated SAP navigation and data entry.

The automation was developed to standardize these repetitive activities and provide a more consistent workflow for Purchasing users.

---

# Solution

The solution uses Excel VBA as the automation controller and SAP GUI Scripting to automate SAP interactions.

The automation supports three functional areas.

### Plan Order / SA Validation

The validation process checks:

- Plan Order availability.
- Required quantity.
- Whether the material is identified as an SA part.

When an SA part is identified, an additional manual check is required to confirm whether sufficient Plan Order quantity is available before proceeding.

### Purchase Requisition Creation

PR creation is the primary workflow.

The automation:

- Reads the required information from Excel.
- Performs the required validation.
- Processes the relevant SAP information.
- Creates the Purchase Requisition.
- Retrieves the SAP processing result.
- Returns the result to Excel.

### Purchase Requisition Close/Fix

A separate workflow is provided for existing Purchase Requisitions that require Close/Fix actions.

The automation:

- Retrieves the existing PR in SAP.
- Performs the required Close/Fix action.
- Saves the PR.
- Retrieves the SAP processing result.
- Updates the Excel result.

# Key Features

- Purchase Requisition creation
- Plan Order quantity validation
- Sales Agreement (SA) part identification
- Manual validation path for SA parts
- SAP GUI automation
- Purchase Requisition Close/Fix
- Excel-based input
- Excel-based result tracking
- SAP processing message handling
- Repeatable Purchasing workflow

---

# Technologies Used

| Category | Technology |
|---|---|
| Programming Language | Excel VBA |
| ERP | SAP MM |
| SAP Automation | SAP GUI Scripting |
| User Interface | Microsoft Excel |
| Data Processing | Excel VBA / Excel formulas |
| Input / Output | Microsoft Excel |

---

# Workflow

The overall automation workflow is shown below.

![SAP PR Automation Workflow](docs/images/workflow.png)

The primary workflow starts from Excel input, performs the required Plan Order / Sales Agreement validation, processes the required SAP information, creates the Purchase Requisition, and returns the SAP processing result to Excel.

The Close/Fix PR process is handled as a separate workflow.

---

# Validation Logic

Plan Order / SA validation supports the Purchase Requisition creation process.

The validation focuses on two conditions:

### Plan Order Availability

The automation checks available Plan Order information against the required quantity.

### Sales Agreement Part

The automation determines whether the material is identified as an SA part.

When an SA part is identified, an additional manual check is required to confirm whether sufficient Plan Order quantity is available.

The validation can result in conditions such as:

```text
OK
Check if SA part
Not enough Plan Order
```

The validation result supports the decision to proceed with PR creation.

---

# SAP GUI Automation

SAP GUI Scripting is used to automate the interaction between the Excel VBA controller and SAP.

The automation performs the required SAP interactions, including:

- Sending data to SAP.
- Navigating SAP screens.
- Processing purchasing information.
- Executing the required actions.
- Reading SAP processing messages.
- Returning results to Excel.

# Architecture

The solution uses Excel as the user-facing layer, VBA as the automation controller, and SAP GUI Scripting as the integration layer with SAP.

![SAP PR Automation Architecture](docs/images/architecture.png)

---

## Architecture Components

| Component | Responsibility |
|---|---|
| Excel Workbook | User input, configuration, tracking and results |
| VBA Controller | Controls the automation workflow |
| Plan Order Validation | Checks Plan Order quantity and SA conditions |
| Create PR Module | Executes the primary PR creation workflow |
| Close/Fix PR Module | Executes the separate PR Close/Fix workflow |
| SAP GUI Scripting | Automates SAP interactions |
| SAP System | Executes SAP processes and returns results |
| Excel Output | Displays processing status, messages and PR results |

---

# Purchase Requisition Creation

Purchase Requisition creation is the **primary function** of this project.

The automation uses Excel input and SAP processing to create the required PR.

The process includes:

1. Reading the required information from Excel.
2. Performing the required Plan Order / SA validation.
3. Processing the required SAP information.
4. Creating the Purchase Requisition.
5. Retrieving the SAP processing result.
6. Updating the Excel result.

The result may include:

- SAP processing message.
- PR number when successfully created.
- Processing status or remarks.

---

# Purchase Requisition Close/Fix

Close/Fix is a **separate workflow** from PR creation.

It is used when an existing Purchase Requisition requires the applicable Close/Fix action.

The automation:

1. Retrieves the existing PR.
2. Opens the PR in SAP.
3. Performs the required Close/Fix action.
4. Saves the PR.
5. Retrieves the SAP processing result.
6. Updates the Excel result.

---

# Excel Integration

Excel acts as the user-facing controller for the automation.

The workbook is used to:

- Provide input data.
- Start the automation.
- Store validation information.
- Track processing results.
- Display SAP messages.
- Display PR numbers and processing status.

This allows Purchasing users to interact with the automation through an environment already familiar to them.

# Screenshots and Results

The repository contains screenshots demonstrating the main automation workflows and their results.

---

## Plan Order Validation

![Plan Order Validation](docs/images/plan-order-validation.png)

This screenshot demonstrates the Plan Order / SA validation process and the information used to support PR creation.

---

## PR Creation Result

![PR Creation Result](docs/images/pr-creation-result.png)

This screenshot demonstrates the result returned after the PR creation process.

---

## Close PR

![Close PR](docs/images/close-pr.png)

This screenshot demonstrates the Close/Fix PR workflow.

---

## Close PR Result

![Close PR Result](docs/images/close-pr-result.png)

This screenshot demonstrates the processing result returned after the Close/Fix workflow.

---

# Business Rules

The automation incorporates Purchasing business rules related to:

### Plan Order Availability

Plan Order information is checked before proceeding with PR creation.

### Sales Agreement Parts

The automation identifies whether the material is an SA part.

When an SA part is identified, an additional manual check is required to confirm sufficient Plan Order availability.

### PR Creation

PR creation proceeds according to the applicable validation and SAP processing conditions.

### PR Close/Fix

Existing PRs can be processed through the separate Close/Fix workflow.

---

# Business Impact

The automation helps the Purchasing team by:

- Reducing repetitive SAP navigation.
- Reducing repetitive manual data entry.
- Standardizing the PR creation workflow.
- Providing a repeatable Plan Order / SA validation process.
- Supporting consistent handling of SA-related conditions.
- Automating repetitive PR Close/Fix activities.
- Returning SAP processing results directly to Excel.
- Allowing users to review processing status within the existing Excel workflow.

No percentage-based efficiency improvement is claimed because formally measured performance data is not available.

# My Role

I was responsible for developing and maintaining this automation for the Purchasing team.

My responsibilities included:

- Understanding the Purchase Requisition business process.
- Identifying repetitive SAP activities suitable for automation.
- Designing the automation workflow.
- Developing the Excel VBA automation.
- Integrating Excel VBA with SAP GUI Scripting.
- Implementing Plan Order validation logic.
- Implementing Sales Agreement part checks.
- Developing the PR creation workflow.
- Developing the separate PR Close/Fix workflow.
- Handling SAP processing results.
- Updating Excel with validation and processing results.
- Testing the automation against the Purchasing workflow.
- Maintaining and enhancing the solution after deployment.

---

# Current Limitations

- Requires SAP GUI.
- Requires appropriate SAP user authorization.
- Requires Microsoft Excel with VBA support.
- Depends on SAP GUI screen elements and scripting identifiers.
- Depends on the existing Excel workbook structure.
- Changes to SAP screens may require maintenance.
- Changes to Purchasing business rules may require updates to the automation.
- The automation is designed around the current Purchasing workflow.

---

## Known Issues / SAP Upgrade Compatibility

### Plan Order / SA Validation

The Plan Order / Sales Agreement (SA) validation workflow was originally developed and tested against the previous SAP system environment.

Following the SAP system upgrade, the underlying SAP SQ00 query:

`DEMO_SAP_QUERY`

is no longer consistently returning the expected report data.

During testing in the upgraded SAP environment, the query can return:

`No data was selected`

This occurs during the SAP report-generation stage, before the exported report is transferred back to Excel.

### Current Status

| Function | Status |
|---|---|
| Purchase Requisition Creation | Working |
| Purchase Requisition Close/Fix | Working |
| Plan Order / SA Validation | Pending SAP upgrade compatibility adjustment |

### Impact

The SAP upgrade affects the Plan Order / SA validation portion of the automation.

The issue is related to the behaviour and/or configuration of the SAP query after the system upgrade rather than the Excel result-processing logic.

The validation workflow is therefore currently considered **SAP environment dependent** and requires further adjustment against the upgraded SAP environment.

### Future Action

The validation workflow can be updated once the upgraded SAP query behaviour is confirmed.

Potential investigation areas include:

- Changes to the `PU-088MRP_EFFI` SQ00 query.
- Changes to query selection criteria.
- Changes to SAP master data or report logic.
- Changes to MRP-related fields after the SAP upgrade.
- Changes to SAP GUI behaviour or screen elements.
- Changes to the report output generated by SQ00.

# Lessons Learned

Developing this automation provided practical experience in automating a real SAP Purchasing process.

Key lessons include:

- Business-process understanding is essential before automating SAP activities.
- Business rules should be understood before automating transaction execution.
- Validation should be separated from execution where appropriate.
- SAP GUI automation requires careful handling of SAP screen elements.
- Excel can provide a practical interface for business users.
- SAP processing results should be captured and returned to users.
- Separate workflows should remain clearly defined when they serve different business purposes.
- Automation should focus on reducing repetitive work while preserving the existing business process.

# Future Improvements

Potential future improvements include:

### Python Migration

Rebuild selected components using Python to improve modularity and maintainability.

### Reusable SAP Modules

Create reusable components for:

- SAP session handling.
- Material processing.
- Plan Order processing.
- PR creation.
- PR Close/Fix.

### Centralized Configuration

Move configurable values and SAP element identifiers into centralized configuration files.

### Structured Logging

Introduce structured logging for:

- Processing start and completion.
- Material information.
- PR numbers.
- SAP messages.
- Errors.

### Improved Validation Layer

Separate business-rule validation from SAP transaction execution to make the logic easier to maintain and test.

### Improved Error Handling

Introduce structured exception handling and recovery for SAP GUI automation failures.

---

# Project Structure

```text
sap-pr-automation/
│
├── Combine PR.xlsm
├── Create_PR.bas
├── Validate_PR.bas
├── Close_Fixed_PR.bas
├── README.md
│
└── docs/
    └── images/
        ├── architecture.png
        ├── close-pr-result.png
        ├── close-pr.png
        ├── plan-order-validation.png
        ├── pr-creation-result.png
        └── workflow.png
```

---

# Engineering Skills Demonstrated

- Excel VBA
- SAP GUI Scripting
- SAP MM
- SAP Purchasing
- Purchase Requisition Automation
- Plan Order Validation
- Sales Agreement Business Rules
- Excel Automation
- Business Process Automation
- SAP Workflow Automation
- Business Rule Implementation
- Error Handling
- SAP Result Processing
- Troubleshooting
- Production Automation Maintenance

---

# Project Information

| Item | Details |
|---|---|
| Project Status | Production / SAP Upgrade Compatibility Review |
| Project Type | SAP Business Process Automation |
| Primary Function | Purchase Requisition Creation |
| Validation | Plan Order / Sales Agreement Pre-Check |
| Separate Function | Purchase Requisition Close/Fix |
| Programming Language | Excel VBA |
| ERP | SAP MM |
| SAP Automation | SAP GUI Scripting |
| User Interface | Microsoft Excel |
| Primary Users | Purchasing Team |

---

# Summary

The solution demonstrates a production Purchasing automation workflow developed using Excel VBA and SAP GUI Scripting.

The primary Purchase Requisition creation and Close/Fix workflows remain functional. The Plan Order / Sales Agreement validation workflow is currently undergoing compatibility review following an SAP system upgrade that affected the underlying SQ00 report generation.

The **primary purpose** of the automation is Purchase Requisition creation.

Plan Order and Sales Agreement validation supports the PR creation process by helping users determine whether the required conditions are met before proceeding.

A **separate Close/Fix workflow** is also provided for processing existing Purchase Requisitions.

The solution combines:

```text
Excel Workbook
      │
      ▼
VBA Controller
      │
      ├── Plan Order / SA Validation
      │
      ├── Create PR
      │
      └── Close/Fix PR
              │
              ▼
       SAP GUI Scripting
              │
              ▼
          SAP System
              │
              ▼
       Processing Result
              │
              ▼
        Excel Workbook
```

This project demonstrates practical experience in understanding a Purchasing business process, implementing SAP GUI automation, applying business rules, and maintaining a production automation solution.
