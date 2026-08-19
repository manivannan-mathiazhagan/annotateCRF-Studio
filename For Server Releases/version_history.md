# Version History

## Version 2.0 – CDISC Library Integration and SUPP Annotation

### CDISC Library Integration

- Integrated AnnotateCRF Studio with the **CDISC Library API**.
- Added support for **SDTMIG 3.2, SDTMIG 3.3 and SDTMIG 3.4**.
- Added automatic retrieval of SDTM domain metadata.
- Added automatic retrieval of SDTM dataset labels.
- Added automatic retrieval of variables belonging to the selected domain.
- Reduced manual maintenance of domain and variable lists.
- Reduced manual typing and annotation inconsistencies.
- Added support for primary and secondary CDISC Library API keys.
- Added validation and error handling when metadata cannot be retrieved.

### SUPP Annotation Support

- Added a dedicated **SUPP Annotation** type.
- Added automatic Supplemental Qualifier domain naming.
- Added annotation text such as `CMAESPID in SUPPCM`.
- Preserved the parent SDTM domain in annotation metadata.
- Allowed manually entered domains and variables where required.

### Additional Improvements

- Improved stable domain colour handling.
- Improved annotation box sizing for long and multi-line text.
- Improved annotation rendering consistency between preview and final PDF.
- Improved Refer Page annotation hyperlinks.
- Improved final PDF output accuracy.

## Version 1.2 – Usability, Reliability and Deployment

- Added multiple annotation selection and delete functionality.
- Added support for deleting multiple selected annotations in one operation.
- Fixed CSV export issues seen in certain cases.
- Improved CSV export reliability based on feedback from Yi-Hsuang.
- Added a standalone portable executable version.
- Enabled users to run the application without a separate Python installation.
- Added support for portrait and landscape pages within the same PDF.
- Improved annotation alignment for mixed-orientation documents.

## Version 1.1 – Navigation and Annotation Productivity

- Added a page number input field with a **Go** button.
- Added direct navigation to a specified CRF page.
- Added multiple annotation selection from the annotation grid.
- Added **Copy Annotation To Page** for copying selected annotations together.
- Improved efficiency when repeated annotations are required across multiple pages.

## Version 1.0 – Initial Release

- Initial release of **AnnotateCRF Studio**.
- Added interactive PDF annotation placement.
- Added Domain, Variable, Assigned Field and NOTSUB annotations.
- Added Refer Page annotations with clickable hyperlinks.
- Added connector line creation and management.
- Added bookmark creation, editing and hierarchy management.
- Added automatic Table of Contents generation.
- Added annotation and bookmark CSV import and export.
- Added review mode.
- Added final annotated PDF generation.
