
{.deadCodeElim: on.}


proc xlsxioread_get_version*(pmajor: ptr cint; pminor: ptr cint;
    pmicro: ptr cint) {.
    cdecl, importc: "xlsxioread_get_version".}


proc xlsxioread_get_version_string*(): cstring {.cdecl,
    importc: "xlsxioread_get_version_string".}

type
  xlsxio_read_struct* = object
  XlsxioReader* = ptr xlsxio_read_struct


proc xlsxioread_open*(filename: cstring): XlsxioReader {.cdecl,
    importc: "xlsxioread_open".}

proc xlsxioread_open_filehandle*(filehandle: cint): XlsxioReader {.cdecl,
    importc: "xlsxioread_open_filehandle".}


proc xlsxioread_open_memory*(data: pointer; datalen: uint64;
    freedata: cint): XlsxioReader {.
    cdecl, importc: "xlsxioread_open_memory".}

proc xlsxioread_close*(handle: XlsxioReader) {.cdecl,
    importc: "xlsxioread_close".}

type
  xlsxioread_list_sheets_callback_fn* = proc (name: cstring;
      callbackdata: pointer): cint {.cdecl.}


proc xlsxioread_list_sheets*(handle: XlsxioReader;
                            callback: xlsxioread_list_sheets_callback_fn;
                            callbackdata: pointer) {.cdecl,
    importc: "xlsxioread_list_sheets".}

const
  XLSXIOREAD_SKIP_NONE* = 0
  XLSXIOREAD_SKIP_EMPTY_ROWS* = 0x01
  XLSXIOREAD_SKIP_EMPTY_CELLS* = 0x02
  XLSXIOREAD_SKIP_ALL_EMPTY* = (
    XLSXIOREAD_SKIP_EMPTY_ROWS or XLSXIOREAD_SKIP_EMPTY_CELLS)
  XLSXIOREAD_SKIP_EXTRA_CELLS* = 0x04
  XLSXIOREAD_SKIP_HIDDEN_ROWS* = 0x08


type
  xlsxioread_process_cell_callback_fn* = proc (row: uint32; col: uint32;
      value: cstring; callbackdata: pointer): cint {.cdecl.}


type
  xlsxioread_process_row_callback_fn* = proc (row: uint32; maxcol: uint32;
      callbackdata: pointer): cint {.cdecl.}



proc xlsxioread_process*(handle: XlsxioReader; sheetname: cstring;
                        flags: cuint;
                        cell_callback: xlsxioread_process_cell_callback_fn;
                        row_callback: xlsxioread_process_row_callback_fn;
                        callbackdata: pointer): cint {.cdecl,
    importc: "xlsxioread_process".}


type
  xlsxio_read_sheetlist_struct* = object
  XlsxioReaderSheetList* = ptr xlsxio_read_sheetlist_struct

proc xlsxioread_sheetlist_open*(handle: XlsxioReader): XlsxioReaderSheetList {.
    cdecl, importc: "xlsxioread_sheetlist_open".}

proc xlsxioread_sheetlist_close*(sheetlisthandle: XlsxioReaderSheetList) {.cdecl,
    importc: "xlsxioread_sheetlist_close".}


proc xlsxioread_sheetlist_next*(sheetlisthandle: XlsxioReaderSheetList): cstring {.
    cdecl, importc: "xlsxioread_sheetlist_next".}

type
  xlsxio_read_sheet_struct* = object
  XlsxioReaderSheet* = ptr xlsxio_read_sheet_struct


proc xlsxioread_sheet_last_row_index*(sheethandle: XlsxioReaderSheet): csize_t {.
    cdecl, importc: "xlsxioread_sheet_last_row_index".}

proc xlsxioread_sheet_last_column_index*(
  sheethandle: XlsxioReaderSheet): csize_t {.
    cdecl, importc: "xlsxioread_sheet_last_column_index".}

proc xlsxioread_sheet_flags*(sheethandle: XlsxioReaderSheet): cuint {.cdecl,
    importc: "xlsxioread_sheet_flags".}


proc xlsxioread_sheet_open*(handle: XlsxioReader; sheetname: cstring;
                           flags: cuint): XlsxioReaderSheet {.cdecl,
    importc: "xlsxioread_sheet_open".}


proc xlsxioread_sheet_close*(sheethandle: XlsxioReaderSheet) {.cdecl,
    importc: "xlsxioread_sheet_close".}


proc xlsxioread_sheet_next_row*(sheethandle: XlsxioReaderSheet): cint {.cdecl,
    importc: "xlsxioread_sheet_next_row".}


proc xlsxioread_sheet_next_cell*(sheethandle: XlsxioReaderSheet): cstring {.
    cdecl, importc: "xlsxioread_sheet_next_cell".}

proc xlsxioread_sheet_next_cell_string*(sheethandle: XlsxioReaderSheet;
                                       pvalue: ptr cstring): cint {.cdecl,
    importc: "xlsxioread_sheet_next_cell_string".}


proc xlsxioread_sheet_next_cell_int*(sheethandle: XlsxioReaderSheet;
                                    pvalue: ptr int64): cint {.cdecl,
    importc: "xlsxioread_sheet_next_cell_int".}

proc xlsxioread_sheet_next_cell_float*(sheethandle: XlsxioReaderSheet;
                                      pvalue: ptr cdouble): cint {.cdecl,
    importc: "xlsxioread_sheet_next_cell_float".}


proc xlsxioread_sheet_next_cell_datetime*(sheethandle: XlsxioReaderSheet;
    pvalue: ptr int64): cint {.cdecl,
                            importc: "xlsxioread_sheet_next_cell_datetime".}

proc xlsxioread_free*(data: cstring) {.cdecl, importc: "xlsxioread_free".}

