# Invoice Generator

Utility to automatize template filling for invoice generation

## How to use

You'll have to pass the values to be replaced: invoiceConsecutive, payment. Values like dateOfIssue and dueDate are automatically generated from the current date.

```
node ./src/index.js --template=path/to/template.docx --output=path/to/output.docx --values="invoiceConsecutive=003,payment=1845"
```
