*"* use this source file for any type of declarations (class
*"* definitions, interfaces or type declarations) you need for
*"* components in the private section
CLASS lcl_file_handler DEFINITION
  CREATE PUBLIC
  FINAL.

  PUBLIC SECTION.
    METHODS select_excel_file
      RETURNING VALUE(result) TYPE REF TO lcl_file_handler.
    METHODS get_file_name
      RETURNING VALUE(result) TYPE string.
    METHODS save_excel_file
      IMPORTING !workbook_content TYPE xstring
                !file_name        TYPE csequence OPTIONAL.

  PRIVATE SECTION.
    DATA file_name TYPE string.
ENDCLASS.
