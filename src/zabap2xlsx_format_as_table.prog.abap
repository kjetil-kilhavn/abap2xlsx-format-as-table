*&---------------------------------------------------------------------*
*& Report ZABAP2XLSX_FORMAT_AS_TABLE
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zabap2xlsx_format_as_table.

LOAD-OF-PROGRAM.
  DATA(application) = NEW zcl_abap2xlsx_format_as_table( ).

START-OF-SELECTION.
  TRY.
      application->main( |Demonstrate issue: 'Format as table' formatting lost| ).
    CATCH BEFORE UNWIND cx_root INTO DATA(error).
      WRITE: / error->get_text( ).
      IF error->is_resumable = abap_true.
        RESUME.
      ENDIF.
  ENDTRY.
