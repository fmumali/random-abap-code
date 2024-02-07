  METHOD resolve_path.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Determine whether the replacement should be done
*                iterative to allow /../../..   or something alike
*        2do§2   Determine whether /./ has to be supported as well
*        2do§3   Create unit-test for this method
*
*                Please don't just delete these ToDos if they are not
*                needed but leave a comment that states this
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-11
*              - ...
* changes: replaced previous coding by regular expression
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* §1  This routine will receive a path, that may have a relative pathname (/../) included somewhere
*     The output should be a resolved path without relative references
*     Example:  Input     xl/worksheets/../drawings/drawing1.xml
*               Output    xl/drawings/drawing1.xml
*--------------------------------------------------------------------*

    rp_result = ip_path.
*--------------------------------------------------------------------*
* §1  Remove relative pathnames
*--------------------------------------------------------------------*
*  Regular expression   [^/]*/\.\./
*                       [^/]*            --> any number of characters other than /
*   followed by              /\.\./      --> the sequence /../
*   ==> worksheets/../ will be found in the example
*--------------------------------------------------------------------*
    REPLACE REGEX '[^/]*/\.\./' IN rp_result WITH ``.


  ENDMETHOD.