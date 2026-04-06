# Future Additions (in no particular order):
- gui
- cad files
- input validation - verify the tag name column contains . for discrete and [] for analog
- module presets?? - not sure about this
- update-project command? - what would this do?
- package into wheel for distribution
- generate hardware?? - is this a bad idea?
- work for other types of I/O (flex, compact) - may already work, needs testing
- readme update with instructions on usage
- different io types
    - 1734 - point io - done
    - 1794 - flex io - done
    - 1756 - control logix io - done
# Completed:
- output directory option (--output) -DONE
- safety program generator - DONE
- put on github - DONE
- auto generate tag name using module type and routine name number - DONE

# Issues:
- IO Module Status Program - GSV instance name uses module name from sheet, but i dont always name hardware the same, for example, GBP C10 has hardware module "R4103b_SIL1" but name in routine is "R4103b_SIL1_IR2"


