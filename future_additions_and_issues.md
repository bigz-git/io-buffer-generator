# Future Additions (in no particular order):
- gui
- cad files
- module presets?? - not sure about this
- update-project command? - what would this do?
- package into wheel for distribution
- generate hardware?? - is this a bad idea?
- readme update with instructions on usage

- --dry-run flag on generate - print what files would be written and their tag/routine counts without writing anything
- fill-drawings command - same pattern as fill-descriptions; prompts for a drawing filename and fills blank column G cells for a selected rack (currently left as "ENTER DRAWING NAME HERE" placeholder)
- summary command / enhanced list - show counts of filled vs. blank tags and descriptions per rack so you can see which racks are ready to generate vs. still incomplete
- remove-rack command - remove a rack sheet and its Cover Sheet entry without opening Excel - DONE

# Completed:
- output directory option (--output) -DONE
- safety program generator - DONE
- put on github - DONE
- auto generate tag name using module type and routine name number - DONE
- different io types
    - 1734 - point io - done
    - 1794 - flex io - done
    - 1756 - control logix io - done
- input validation - verify the tag name column contains . for discrete and [] for analog - DONE (additional error checking too)
- validate command - run all error-checking from read_project without generating files; lets you catch tag/routine name errors incrementally as you fill in the sheet instead of only at generate time - DONE
- rename-rack command - rename a rack sheet and update the Cover Sheet summary row to match, without opening Excel - DONE

# Issues:
- IO Module Status Program - GSV instance name uses module name from sheet, but i dont always name hardware the same, for example, GBP C10 has hardware module "R4103b_SIL1" but name in routine is "R4103b_SIL1_IR2"


