
{.deadCodeElim: on.}

proc xlsxiowrite_get_version*(pmajor: ptr cint; pminor: ptr cint;
    pmicro: ptr cint) {.
    cdecl, importc: "xlsxiowrite_get_version".}

proc xlsxiowrite_get_version_string*(): cstring {.
    cdecl, importc: "xlsxiowrite_get_version_string".}

type
  Xlsxio_write_struct* = object
  Xlsxiowriter* = ptr Xlsxio_write_struct



proc xlsxiowrite_open*(filename: cstring; sheetname: cstring): Xlsxiowriter {.
    cdecl, importc: "xlsxiowrite_open".}


proc xlsxiowrite_close*(handle: Xlsxiowriter): cint {.cdecl,
    importc: "xlsxiowrite_close".}


proc xlsxiowrite_set_detection_rows*(handle: Xlsxiowriter; rows: csize_t) {.
    cdecl, importc: "xlsxiowrite_set_detection_rows".}

proc xlsxiowrite_set_row_height*(handle: Xlsxiowriter; height: csize_t) {.
    cdecl, importc: "xlsxiowrite_set_row_height".}

proc xlsxiowrite_add_column*(handle: Xlsxiowriter; name: cstring;
    width: cint) {.
    cdecl, importc: "xlsxiowrite_add_column".}

proc xlsxiowrite_add_cell_string*(handle: Xlsxiowriter; value: cstring) {.
    cdecl, importc: "xlsxiowrite_add_cell_string".}

proc xlsxiowrite_add_cell_int*(handle: Xlsxiowriter; value: int64) {.
    cdecl, importc: "xlsxiowrite_add_cell_int".}

proc xlsxiowrite_add_cell_float*(handle: Xlsxiowriter; value: cdouble) {.
    cdecl, importc: "xlsxiowrite_add_cell_float".}


proc xlsxiowrite_add_cell_datetime*(handle: Xlsxiowriter; value: int64) {.
    cdecl, importc: "xlsxiowrite_add_cell_datetime".}

proc xlsxiowrite_next_row*(handle: Xlsxiowriter) {.cdecl,
    importc: "xlsxiowrite_next_row".}

proc xlsxioread_free(data: cstring){.cdecl,
    importc: "xlsxiowrite_next_row".}