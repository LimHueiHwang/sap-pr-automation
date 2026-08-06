# SAP Purchase Requisition Automation

![Production](https://img.shields.io/badge/Status-Production-success)
![SAP MM](https://img.shields.io/badge/SAP-MM-blue)
![Excel VBA](https://img.shields.io/badge/VBA-Excel-green)

> Internal SAP GUI automation developed to validate, create, and remove Purchase Requisitions (PR) for the Purchasing team.

## 📌 Project Summary

This project automates the validation, creation, and removal of SAP Purchase Requisitions (PR) for the Purchasing team.

It was developed to reduce repetitive manual SAP processing while improving consistency through automated business rule validation.

The solution integrates Microsoft Excel VBA with SAP GUI Scripting to streamline PR processing without changing existing business workflows.

**Project Status:** Production

## 🎯 Business Problem

Before this automation was introduced, buyers manually:

- Validated Plan Orders.
- Verified supplier-specific business rules.
- Created Purchase Requisitions.
- Removed unnecessary Purchase Requisitions.

The manual workflow required repetitive SAP navigation and multiple validation steps, making the process time-consuming and increasing the risk of inconsistent processing.

## 💡 Solution

This automation standardizes the Purchase Requisition workflow by:

- Validating Plan Orders before PR creation.
- Applying predefined business rules.
- Automating Purchase Requisition creation.
- Automating Purchase Requisition removal.
- Reducing repetitive SAP interactions.

The objective is to improve processing consistency while reducing manual effort for buyers.

## ✨ Core Features

### Purchase Requisition Processing

- Validate Plan Orders.
- Create Purchase Requisitions.
- Remove Purchase Requisitions.

### Automation

- SAP GUI Automation.
- Business Rule Validation.
- Excel-based User Interface.

## 🛠 Technologies Used

| Category | Technology |
|----------|------------|
| Language | Excel VBA |
| ERP | SAP MM |
| Automation | SAP GUI Scripting |
| Office | Microsoft Excel |

## 📊 Workflow

The following diagram illustrates the business workflow of the automation.

![Workflow](docs/images/workflow.png)

## 🏗 Architecture

The following diagram illustrates the high-level architecture of the solution.

![Architecture](docs/images/architecture.png)

## 📷 Screenshots

### Main Interface

![Main Interface](docs/images/excel-main.png)

---

### SAP Processing

![SAP Processing](docs/images/sap-processing.png)

---

### Processing Result

![Result](docs/images/result.png)

## 📈 Business Impact

This automation helps the Purchasing team by:

- Reducing repetitive manual SAP processing.
- Standardizing Purchase Requisition validation.
- Improving consistency during PR processing.
- Reducing repetitive SAP navigation for buyers.
- Supporting repeatable purchasing workflows.

## 👨‍💻 My Role

I was responsible for:

- Understanding business requirements.
- Designing the automation workflow.
- Developing the Excel VBA solution.
- Integrating SAP GUI Scripting.
- Testing the automation.
- Maintaining and enhancing the solution after deployment.

## 📋 Business Rules

The automation validates:

- Plan Order availability.
- Supplier-specific requirements.
- Purchase Requisition eligibility.
- SAP processing conditions.

Processing only continues after all validation rules have passed.

## ⚠ Current Limitations

- Requires SAP GUI.
- Requires SAP GUI Scripting to be enabled.
- Requires user authorization.
- Designed specifically for internal Purchasing workflows.

## 📚 Lessons Learned

Developing this project reinforced several important engineering principles:

- Business rules should be validated before automating SAP transactions.
- Separating validation from execution improves maintainability.
- Automation should be designed around business processes rather than individual SAP screens.
- Modular design makes future enhancements easier.

## 🚀 Roadmap

### Version 2

- Migrate the solution to Python.
- Replace Excel-based validation with pandas.
- Separate SAP operations into reusable modules.
- Introduce centralized logging.
- Support parallel SAP sessions.
- Store configuration in external configuration files.

### Long-Term Vision

- Integrate into a centralized Automation Center.
- Improve scalability and maintainability.
- Support additional purchasing automation workflows.

![Roadmap](docs/images/roadmap.png)

## 📦 Version History

### Version 1

- Excel VBA
- SAP GUI Scripting
- Production deployment

### Planned Version 2

- Python implementation
- pandas-based validation
- Parallel SAP session processing
- Reusable automation framework
- Centralized logging
- Modular architecture

## 🛠 Engineering Skills Demonstrated

- SAP GUI Automation
- Business Process Automation
- Excel VBA Development
- Business Rule Validation
- SAP MM
- Workflow Automation
- Process Standardization
- SAP Integration

## 📊 Project Statistics

| Item | Details |
|------|---------|
| Project Status | Production |
| Project Type | Internal SAP Automation |
| ERP System | SAP MM |
| Programming Language | Excel VBA |
| Automation Method | SAP GUI Scripting |
| Primary Users | Purchasing Team |

## 💭 Reflection

This project was originally developed to solve an immediate business need using Excel VBA and SAP GUI Scripting.

Since completing the project, I have gained additional experience in Python automation and software architecture. If redesigning the solution today, I would build it using Python with reusable modules, pandas-based validation, centralized logging, and parallel SAP session processing while preserving the same business objectives.

This project represents an important step in my progression from VBA-based automation towar
