*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations
CLASS lcl_file_handler IMPLEMENTATION.

  METHOD select_excel_file.
    DATA selected_files TYPE filetable.
    DATA return_code TYPE i.
    DATA user_action TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING
        file_filter    = |{ cl_gui_frontend_services=>filetype_excel }{ cl_gui_frontend_services=>filetype_all }|
        multiselection = abap_false
      CHANGING
        file_table     = selected_files
        rc             = return_code
        user_action    = user_action
      EXCEPTIONS
        file_open_dialog_failed = 1
        cntl_error              = 2
        error_no_gui            = 3
        not_supported_by_gui    = 4 ).
    ASSERT sy-subrc = 0.
    IF user_action = cl_gui_frontend_services=>action_cancel.
      RETURN.
    ENDIF.
    ASSERT user_action = cl_gui_frontend_services=>action_ok.
    file_name = selected_files[ 1 ]-filename.

    result = me.
  ENDMETHOD.

  METHOD get_file_name.
    result = file_name.
  ENDMETHOD.

  METHOD save_excel_file.
    CHECK workbook_content IS NOT INITIAL.

    DATA(new_file_name) = COND string( WHEN file_name IS INITIAL
                                         THEN me->file_name
                                       ELSE file_name ).
    DATA(file_size) = xstrlen( workbook_content ).
    DATA(file_content) = cl_bcs_convert=>xstring_to_solix( workbook_content ).
    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize              = file_size
        filename                  = new_file_name
        filetype                  = 'BIN'
      CHANGING
        data_tab                  = file_content
      EXCEPTIONS
        file_write_error          = 1
        no_batch                  = 2
        gui_refuse_filetransfer   = 3
        invalid_type              = 4
        no_authority              = 5
        unknown_error             = 6
        header_not_allowed        = 7
        separator_not_allowed     = 8
        filesize_not_allowed      = 9
        header_too_long           = 10
        dp_error_create           = 11
        dp_error_send             = 12
        dp_error_write            = 13
        unknown_dp_error          = 14
        access_denied             = 15
        dp_out_of_memory          = 16
        disk_full                 = 17
        dp_timeout                = 18
        file_not_found            = 19
        dataprovider_exception    = 20
        control_flush_error       = 21
        not_supported_by_gui      = 22
        error_no_gui              = 23 ).
    ASSERT sy-subrc = 0.
  ENDMETHOD.

ENDCLASS.
