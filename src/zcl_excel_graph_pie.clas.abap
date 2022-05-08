CLASS zcl_excel_graph_pie DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_graph
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_GRAPH_PIE
*"* do not include other source files here!!!
    DATA ns_legendposval TYPE string VALUE 'r'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_overlayval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_pprrtl TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_endpararprlang TYPE string VALUE 'it-IT'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_varycolorsval TYPE string VALUE '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_firstsliceangval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showlegendkeyval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showvalval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showcatnameval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showsernameval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showpercentval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showbubblesizeval TYPE string VALUE '0'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
    DATA ns_showleaderlinesval TYPE string VALUE '1'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .

    METHODS clone REDEFINITION.
    METHODS set_show_legend_key
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_values
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_cat_name
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_ser_name
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_percent
      IMPORTING
        !ip_value TYPE c .
    METHODS set_show_leader_lines
      IMPORTING
        !ip_value TYPE c .
    METHODS set_varycolor
      IMPORTING
        !ip_value TYPE c .
  PROTECTED SECTION.
*"* protected components of class ZCL_EXCEL_GRAPH_PIE
*"* do not include other source files here!!!
    METHODS clone_attributes_to REDEFINITION.
  PRIVATE SECTION.
*"* private components of class ZCL_EXCEL_GRAPH_PIE
*"* do not include other source files here!!!
ENDCLASS.



CLASS ZCL_EXCEL_GRAPH_PIE IMPLEMENTATION.


  METHOD clone.
    DATA: lo_pie TYPE REF TO zcl_excel_graph_pie.

    CREATE OBJECT lo_pie TYPE zcl_excel_graph_pie.
    clone_attributes_to( lo_pie ).

    ro_object = lo_pie.
  ENDMETHOD.


  METHOD clone_attributes_to.
    DATA: lo_pie TYPE REF TO zcl_excel_graph_pie.

    super->clone_attributes_to( io_excel_graph ).

    lo_pie ?= io_excel_graph.
    lo_pie->ns_legendposval       = ns_legendposval.
    lo_pie->ns_overlayval         = ns_overlayval.
    lo_pie->ns_pprrtl             = ns_pprrtl.
    lo_pie->ns_endpararprlang     = ns_endpararprlang.
    lo_pie->ns_varycolorsval      = ns_varycolorsval.
    lo_pie->ns_firstsliceangval   = ns_firstsliceangval.
    lo_pie->ns_showlegendkeyval   = ns_showlegendkeyval.
    lo_pie->ns_showvalval         = ns_showvalval.
    lo_pie->ns_showcatnameval     = ns_showcatnameval.
    lo_pie->ns_showsernameval     = ns_showsernameval.
    lo_pie->ns_showpercentval     = ns_showpercentval.
    lo_pie->ns_showbubblesizeval  = ns_showbubblesizeval.
    lo_pie->ns_showleaderlinesval = ns_showleaderlinesval.

  ENDMETHOD.


  METHOD set_show_cat_name.
    ns_showcatnameval = ip_value.
  ENDMETHOD.


  METHOD set_show_leader_lines.
    ns_showleaderlinesval = ip_value.
  ENDMETHOD.


  METHOD set_show_legend_key.
    ns_showlegendkeyval = ip_value.
  ENDMETHOD.


  METHOD set_show_percent.
    ns_showpercentval = ip_value.
  ENDMETHOD.


  METHOD set_show_ser_name.
    ns_showsernameval = ip_value.
  ENDMETHOD.


  METHOD set_show_values.
    ns_showvalval = ip_value.
  ENDMETHOD.


  METHOD set_varycolor.
    ns_varycolorsval = ip_value.
  ENDMETHOD.
ENDCLASS.
