"""
Generates Studio 5000 .l5x import files from a Project model.

Produces up to two files:
  IO_Files_Rev_<mmddyy_hhmm>.l5x        — standard IO buffer + module status programs
  Safety_IO_Files_Rev_<mmddyy_hhmm>.l5x — safety IO buffer + module status programs (if safety modules present)
"""

import os
import warnings
from datetime import datetime

from .models import Project, Rack, Module, Bit, DIGITAL_TYPES, ANALOG_TYPES, SAFETY_TYPES, OTHER_TYPES, IO_FAMILY_FLEX, IO_FAMILY_CLX, IO_FAMILY_FLEX5000


# ---------------------------------------------------------------------------
# UDT XML (static, matches VBA setupDataTypes exactly)
# ---------------------------------------------------------------------------

_UDT_QP_PLC_TAGS_v02 = (
    '<DataType Name="QP_PLC_TAGS_v02" Family="NoFamily" Class="User"><Members>\n'
    '<Member Name="ZZZZZZZZZZQP_PLC_TAG0" DataType="SINT" Dimension="0" Radix="Decimal" Hidden="true" ExternalAccess="Read/Write"/>\n'
    '<Member Name="_Internal_ES_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="0" ExternalAccess="Read/Write"><Description><![CDATA[PLC E-Stop Reset Internal Toggle Bit]]></Description></Member>\n'
    '<Member Name="_Internal_Fault_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="1" ExternalAccess="Read/Write"><Description><![CDATA[PLC Fault Reset Internal Toggle Bit]]></Description></Member>\n'
    '<Member Name="_Always_Off" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="2" ExternalAccess="Read/Write"><Description><![CDATA[Always Off]]></Description></Member>\n'
    '<Member Name="_Always_On" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="3" ExternalAccess="Read/Write"><Description><![CDATA[Always On]]></Description></Member>'
    '<Member Name="_First_Scan" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="4" ExternalAccess="Read/Write"><Description><![CDATA[First Scan Bit]]></Description></Member>'
    '<Member Name="_P_Faults_Detect" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="5" ExternalAccess="Read/Write"><Description><![CDATA[Faults Detect Permissive]]></Description></Member>\n'
    '<Member Name="_P_EStop_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="6" ExternalAccess="Read/Write"><Description><![CDATA[Main Task E-Stop Reset Permissive]]></Description></Member>'
    '<Member Name="_P_Modules_OK" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG0" BitNumber="7" ExternalAccess="Read/Write"><Description><![CDATA[PLC IO Modules No Faults Detected E-Stop Permissive]]></Description></Member>'
    '<Member Name="ZZZZZZZZZZQP_PLC_TAG9" DataType="SINT" Dimension="0" Radix="Decimal" Hidden="true" ExternalAccess="Read/Write"/>\n'
    '<Member Name="_P_Module_Faults_Detect" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="0" ExternalAccess="Read/Write"><Description><![CDATA[Module Faults Detect Permissive]]></Description></Member>'
    '<Member Name="_R_Module_Faults_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="1" ExternalAccess="Read/Write"><Description><![CDATA[Module Faults Reset Request]]></Description></Member>\n'
    '<Member Name="_P_Modules_Fault_Log" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="2" ExternalAccess="Read/Write"><Description><![CDATA[Module Faults Logging Enable Permissive]]></Description></Member>'
    '<Member Name="_R_Faults_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="3" ExternalAccess="Read/Write"><Description><![CDATA[Faults Reset Request]]></Description></Member>\n'
    '<Member Name="_R_Drives_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="4" ExternalAccess="Read/Write"><Description><![CDATA[Drives Reset Request]]></Description></Member>'
    '<Member Name="_R_Modules_Fault_Mem_Clear" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="5" ExternalAccess="Read/Write"><Description><![CDATA[Module Faults Log Memory Clear Request]]></Description></Member>\n'
    '<Member Name="_S_Modules_OK" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="6" ExternalAccess="Read/Write"><Description><![CDATA[PLC IO Modules All Modules OK]]></Description></Member>'
    '<Member Name="_S_Power_Up_OK" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG9" BitNumber="7" ExternalAccess="Read/Write"><Description><![CDATA[Power Up OK]]></Description></Member>'
    '<Member Name="ZZZZZZZZZZQP_PLC_TAG18" DataType="SINT" Dimension="0" Radix="Decimal" Hidden="true" ExternalAccess="Read/Write"/>\n'
    '<Member Name="_S_Coast_Stop_OK" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="0" ExternalAccess="Read/Write"><Description><![CDATA[Drives Coast Stop Checkbacks OK]]></Description></Member>'
    '<Member Name="_S_Quick_Stop_OK" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="1" ExternalAccess="Read/Write"><Description><![CDATA[Drives Quick Stop Checkbacks OK]]></Description></Member>\n'
    '<Member Name="_S_EStop_OK" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="2" ExternalAccess="Read/Write"><Description><![CDATA[E-Stop Reset OK]]></Description></Member>'
    '<Member Name="_S_Master_Perm" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="3" ExternalAccess="Read/Write"><Description><![CDATA[Drives Master Run Permissive]]></Description></Member>\n'
    '<Member Name="_ONS_Modules_Fault_Log" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="4" ExternalAccess="Read/Write"><Description><![CDATA[Module Faults Logging Enable ONS]]></Description></Member>'
    '<Member Name="_OS_Faults_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="5" ExternalAccess="Read/Write"><Description><![CDATA[Faults Reset One-Shot]]></Description></Member>\n'
    '<Member Name="_OS_Drives_Reset" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="6" ExternalAccess="Read/Write"><Description><![CDATA[Drives Reset One-Shot]]></Description></Member>\n'
    '<Member Name="_S_Flasher_1_On" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG18" BitNumber="7" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 1 On Pulse 1000ms On / 1000ms Off]]></Description></Member>\n'
    '<Member Name="ZZZZZZZZZZQP_PLC_TAG27" DataType="SINT" Dimension="0" Radix="Decimal" Hidden="true" ExternalAccess="Read/Write"/>'
    '<Member Name="_S_Flasher_2_On" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG27" BitNumber="0" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 2 On Pulse 1200ms On / 600ms Off]]></Description></Member>\n'
    '<Member Name="_S_Flasher_3_On" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG27" BitNumber="1" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 3 On Pulse]]></Description></Member>'
    '<Member Name="_C_Modules_Fault_Mem_Clear" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_PLC_TAG27" BitNumber="2" ExternalAccess="Read/Write"><Description><![CDATA[Module Faults Log Memory Clear Command]]></Description></Member>\n'
    '<Member Name="_PaR_Modules_Fault_Timers_Preset" DataType="DINT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Module Fault Timers Preset mSec]]></Description></Member>'
    '<Member Name="_TmR_Module_Fault_Perm" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Modules Fault Detect Permissive Off-Delay Timer]]></Description></Member>\n'
    '<Member Name="_Run_Timer" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Main Task Run Mode On-Delay Timer]]></Description></Member>'
    '<Member Name="_Power_Up_Timer" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Power Up OK On-Delay Timer]]></Description></Member>\n'
    '<Member Name="_TmR_Drives_Reset" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Drives Reset Timer]]></Description></Member>'
    '<Member Name="_TmR_Faults_Reset" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Faults Reset Timer]]></Description></Member>\n'
    '<Member Name="_TmR_Flasher_1_On" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 1 On Timer]]></Description></Member>'
    '<Member Name="_TmR_Flasher_1_Off" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 1 Off Timer]]></Description></Member>'
    '<Member Name="_TmR_Flasher_2_On" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 2 On Timer]]></Description></Member>\n'
    '<Member Name="_TmR_Flasher_2_Off" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 2 Off Timer]]></Description></Member>'
    '<Member Name="_TmR_Flasher_3_On" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 3 On Timer]]></Description></Member>\n'
    '<Member Name="_TmR_Flasher_3_Off" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Flasher 3 Off Timer]]></Description></Member>'
    '<Member Name="_TmR_Lamp_Test" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Lamp Test Timer]]></Description></Member>'
    '</Members></DataType>'
)

_UDT_QP_MODULE_TAGS_v02 = (
    '<DataType Name="QP_MODULE_TAGS_v02" Family="NoFamily" Class="User"><Members>'
    '<Member Name="_S_EntryStatus" DataType="INT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Entry Status]]></Description></Member>'
    '<Member Name="_S_FaultCode" DataType="INT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Module Fault Code]]></Description></Member>'
    '<Member Name="_S_ForceStatus" DataType="INT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Force Status]]></Description></Member>\n'
    '<Member Name="_S_LEDStatus" DataType="INT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[LED Status]]></Description></Member>'
    '<Member Name="_S_Mode" DataType="INT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Mode]]></Description></Member>\n'
    '<Member Name="_S_FaultInfo" DataType="DINT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Fault Info]]></Description></Member>'
    '<Member Name="_S_Instance" DataType="DINT" Dimension="0" Radix="Decimal" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Instance]]></Description></Member>'
    '<Member Name="ZZZZZZZZZZQP_MODULEv7" DataType="SINT" Dimension="0" Radix="Decimal" Hidden="true" ExternalAccess="Read/Write"/>\n'
    '<Member Name="_S_Fault" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_MODULEv7" BitNumber="0" ExternalAccess="Read/Write"><Description><![CDATA[Module Fault]]></Description></Member>'
    '<Member Name="_S_Fault_On_ONS" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_MODULEv7" BitNumber="1" ExternalAccess="Read/Write"><Description><![CDATA[Fault On ONS bit]]></Description></Member>\n'
    '<Member Name="_S_Fault_Off_ONS" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_MODULEv7" BitNumber="2" ExternalAccess="Read/Write"><Description><![CDATA[Fault Off ONS bit]]></Description></Member>\n'
    '<Member Name="_S_CommsOK" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_MODULEv7" BitNumber="3" ExternalAccess="Read/Write"><Description><![CDATA[Module Comms OK]]></Description></Member>\n'
    '<Member Name="_S_CommsFault" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZQP_MODULEv7" BitNumber="4" ExternalAccess="Read/Write"><Description><![CDATA[Module Comms Fault]]></Description></Member>\n'
    '<Member Name="_S_FaultTmr" DataType="TIMER" Dimension="0" Radix="NullType" Hidden="false" ExternalAccess="Read/Write"><Description><![CDATA[Module Fault Timer]]></Description></Member>'
    '</Members></DataType>'
)


# Registry — add every new UDT XML string here to have it included in list_udts()
_ALL_UDT_XML = (
    _UDT_QP_PLC_TAGS_v02,
    _UDT_QP_MODULE_TAGS_v02,
)

# 137 matches the column width of the VBA-generated separator lines.
_SEPARATOR = "=" * 137

# Point IO backplane slot ranges for SlotStatusBits registers.
_POINT_IO_SLOTS_0_31_LIMIT  = 32   # slots 0–31  → SlotStatusBits0_31
_POINT_IO_SLOTS_32_63_LIMIT = 63   # slots 32–62 → SlotStatusBits32_63 (63+ unsupported)

# Instruction names that changed in Studio 5000 v37. Add future renames here.
_V37_RENAMES = {
    "NEQ": "NE",
    "EQU": "EQ",
    "MOV": "MOVE",
}


def _instr(name: str, software_version: str) -> str:
    """Return the correct instruction mnemonic for the given software version string."""
    try:
        major = int(software_version.split(".")[0])
    except (ValueError, IndexError):
        warnings.warn(f"Could not parse software version {software_version!r}; defaulting to pre-v37 instructions.")
        major = 0
    if major >= 37:
        return _V37_RENAMES.get(name, name)
    return name


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

_UDT_TYPES = {"QP_PLC_TAGS_v02", "QP_MODULE_TAGS_v02"}
# BOOL and DINT get Class ="Standard"/"Safety" but no Radix.
# INT arrays get Radix="Decimal" but no Class.
# UDT types get neither.


def _tag_xml(name: str, tag_class: str, datatype: str, description: str,
             dimension: int, comments: list = None) -> str:
    """
    Build a <Tag .../> XML element matching VBA output format exactly.
    dimension: -1 = scalar, >=0 = array dimension
    comments: list of (operand, desc) tuples e.g. [(".0", "desc"), ("[1]", "desc")]
    tag_class: "Standard" or "Safety"
    """
    is_udt   = datatype in _UDT_TYPES
    is_array = (dimension >= 0) and not is_udt
    is_class_type = datatype in {"BOOL", "DINT"}  # get Class attribute

    if is_udt:
        attrs = f'Name="{name}" TagType="Base" DataType="{datatype}" Constant="false" ExternalAccess="Read/Write"'
    elif is_array:
        radix = "Float" if datatype == "REAL" else "Decimal"
        attrs = (f'Name="{name}" TagType="Base" DataType="{datatype}"'
                 f' Dimensions="{dimension}" Radix="{radix}"'
                 f' Constant="false" ExternalAccess="Read/Write"')
    elif is_class_type:
        cls = "Safety" if tag_class == "Safety" else "Standard"
        attrs = (f'Name="{name}" Class ="{cls}" TagType="Base"'
                 f' DataType="{datatype}" Constant="false" ExternalAccess="Read/Write"')
    else:
        # fallback for any other primitive
        attrs = f'Name="{name}" TagType="Base" DataType="{datatype}" Constant="false" ExternalAccess="Read/Write"'

    parts = [f'<Tag {attrs}>']

    if description:
        parts.append(f'<Description>\n<![CDATA[{description}]]>\n</Description>')

    if comments:
        comment_lines = ['<Comments>', '']  # blank line after <Comments>
        for operand, desc in comments:
            if desc:
                comment_lines.append(f'<Comment Operand="{operand}">')
                comment_lines.append(f'<![CDATA[{desc}]]>')
                comment_lines.append('</Comment>')
        parts.append('\n'.join(comment_lines))
        parts.append('</Comments>')

    parts.append('</Tag>')
    return '\n'.join(parts)


def _rung_xml(number: int, comment: str, ladder: str) -> str:
    parts = [f'<Rung Number="{number}" Type="N">']
    if comment:
        parts.append(f'<Comment>\n<![CDATA[{comment}]]>\n</Comment>')
    parts.append(f'<Text>\n<![CDATA[{ladder};]]>\n</Text>')
    parts.append('</Rung>')
    return '\n'.join(parts)


def _routine_xml(name: str, rungs: list) -> str:
    parts = [f'<Routine Name="{name}" Type="RLL">', '<RLLContent>']
    parts.extend(rungs)
    parts.append('</RLLContent>')
    parts.append('</Routine>')
    return '\n'.join(parts)


# ---------------------------------------------------------------------------
# Tag base-name extraction
# ---------------------------------------------------------------------------

def _base_name(tag: str, separator: str) -> str:
    """Strip suffix starting at separator. 'CONV_01_IN.0' → 'CONV_01_IN'"""
    idx = tag.find(separator)
    if idx == -1:
        return tag
    return tag[:idx]


def _tag_operand(tag: str, separator: str) -> str:
    """Return suffix from separator onwards. 'CONV_01_IN.0' → '.0'"""
    idx = tag.find(separator)
    if idx == -1:
        return ""
    return tag[idx:]


def _module_base(routine_name: str) -> str:
    """Strip last _SEGMENT from routine name. 'CONV_01_IN_BUFF' → 'CONV_01_IN'"""
    idx = routine_name.rfind("_")
    if idx == -1:
        return routine_name
    return routine_name[:idx]


# ---------------------------------------------------------------------------
# Determine separator character for a module's tags
# ---------------------------------------------------------------------------

def _separator_char(module: Module) -> str:
    """Return '.' for digital/safety, '[' for analog."""
    if module.type in ANALOG_TYPES:
        return "["
    return "."


# ---------------------------------------------------------------------------
# Controller tag builders
# ---------------------------------------------------------------------------

def _build_ctrl_tags(project: Project):
    """
    Returns (ctrl_tags_xml_list, sfty_ctrl_tags_xml_list).
    ctrl_tags: standard controller tags (PLC, per-rack, per-module)
    sfty_ctrl_tags: safety controller tags (per safety module buffer tag)
    """
    ctrl = []
    sfty = []

    # Generic PLC tag
    ctrl.append(_tag_xml("PLC", "Standard", "QP_PLC_TAGS_v02", "Generic PLC Tags", -1))

    for rack in project.racks:
        # Rack-level QP_MODULE_TAGS_v02
        ctrl.append(_tag_xml(
            rack.name, "Standard", "QP_MODULE_TAGS_v02",
            f"{rack.name}\nMain Enet Module", -1
        ))

        for mod in rack.modules:
            if not mod.routine:
                continue

            sep = _separator_char(mod)
            base = _module_base(mod.routine)
            is_safety = mod.type in SAFETY_TYPES
            dest = sfty if is_safety else ctrl
            tag_class = "Safety" if is_safety else "Standard"

            dest.append(_tag_xml(mod.routine, tag_class, "QP_MODULE_TAGS_v02",
                                 f"{mod.routine}\nModule", -1))

            if mod.type in ANALOG_TYPES:
                # CLX and Flex 5000 hardware exposes analog channels as REAL; others use INT
                analog_dtype = "REAL" if rack.io_family in (IO_FAMILY_CLX, IO_FAMILY_FLEX5000) else "INT"
                comments = [(_tag_operand(b.tag, sep), b.description)
                            for b in mod.bits if b.description]
                ctrl.append(_tag_xml(
                    _base_name(mod.bits[0].tag, sep) if mod.bits else mod.routine,
                    "Standard", analog_dtype, "", len(mod.bits), comments
                ))

            elif is_safety:
                comments = [(_tag_operand(b.tag, sep), b.description)
                            for b in mod.bits if b.description]
                sfty.append(_tag_xml(
                    _base_name(mod.bits[0].tag, sep) if mod.bits else base,
                    "Safety", "DINT", "", -1, comments
                ))

            elif mod.type not in OTHER_TYPES:  # Digital
                comments = [(_tag_operand(b.tag, sep), b.description)
                            for b in mod.bits if b.description]
                ctrl.append(_tag_xml(
                    _base_name(mod.bits[0].tag, sep) if mod.bits else mod.routine,
                    "Standard", "DINT", "", -1, comments
                ))

    return ctrl, sfty


# ---------------------------------------------------------------------------
# Buffer routine builders (IO_Buffer_Files program)
# ---------------------------------------------------------------------------

def _buff_routine_comment(is_safety: bool = False) -> str:
    if is_safety:
        comment = (
            f"{_SEPARATOR}\n"
            f"Safety IO Buffer Subroutine\n"
            f"{_SEPARATOR}\n"
        )
    else:
        comment = (
            f"{_SEPARATOR}\n"
            f"IO Buffer Subroutine\n"
            f"{_SEPARATOR}\n"
            f"This routine has an internal toggle bit to enable or disable the routines JSR instruction in the Subroutine Calls routine.\n"
            f"This bit is intended for commissioning and maintenance functions and must be high for normal line operations.\n"
            f"{_SEPARATOR}\n"
            f"Module Fault Detect Logic -- Master Control Relay\n"
            f"<<  When the MCR is disabled, the rung-condition-in is false for all the instructions inside this subroutine >>\n"
            f"{_SEPARATOR}"
        )

    return comment
    


def _build_buffer_routine(rack: Rack, mod: Module, io_card: str, software_version: str = "36.00") -> str:
    rungs = []
    rung_num = 0
    mov = _instr("MOV", software_version)

    # "Other" modules: blank routine with only the local JSR enable bit on rung 0
    if mod.type in OTHER_TYPES:
        rungs.append(_rung_xml(0, "", f"XIC(JSR_ENABLE_{mod.routine})NOP()"))
        return _routine_xml(mod.routine, rungs)

    # MCR opening rung — safety routines use a simplified form (fault conditions
    # live in the safety controller, not exposed here)
    if mod.type in SAFETY_TYPES:
        mcr_ladder = f"NOP()"
    else:
        mcr_ladder = (
            f"XIC(JSR_ENABLE_{mod.routine})"
            f"XIO({io_card}._S_CommsFault)"
            f"XIO({rack.name}._S_Fault)"
            f"XIO({mod.routine}._S_Fault)"
            f"MCR()"
        )
    rungs.append(_rung_xml(rung_num, _buff_routine_comment(mod.type in SAFETY_TYPES), mcr_ladder))
    rung_num += 1

    for bit in mod.bits:
        tag = bit.tag
        b = bit.index
        # Flex IO (1794): first module is slot 0; all others: first module is slot 1
        slot = mod.slot - 1 if rack.io_family == IO_FAMILY_FLEX else mod.slot

        if mod.type == "Input":
            if rack.io_family == IO_FAMILY_FLEX5000:
                ladder = f"XIC({rack.name}:{slot}:I.Pt{b:02d}.Data)OTE({tag})"
            elif (rack.io_family == IO_FAMILY_CLX
                    or (rack.io_family == IO_FAMILY_FLEX and len(mod.bits) == 32)):
                ladder = f"XIC({rack.name}:{slot}:I.Data.{b})OTE({tag})"
            else:
                ladder = f"XIC({rack.name}:{slot}:I.{b})OTE({tag})"

        elif mod.type == "Output":
            if rack.io_family == IO_FAMILY_FLEX5000:
                ladder = f"XIC({tag})OTE({rack.name}:{slot}:O.Pt{b:02d}.Data)"
            elif (rack.io_family == IO_FAMILY_CLX
                    or (rack.io_family == IO_FAMILY_FLEX and len(mod.bits) == 32)):
                ladder = f"XIC({tag})OTE({rack.name}:{slot}:O.Data.{b})"
            else:
                ladder = f"XIC({tag})OTE({rack.name}:{slot}:O.{b})"

        elif mod.type == "Analog Input":
            if rack.io_family == IO_FAMILY_FLEX5000:
                ladder = f"{mov}({rack.name}:{slot}:I.Ch{b:02d}.Data,{tag})"
            else:
                ladder = f"{mov}({rack.name}:{slot}:I.Ch{b}Data,{tag})"

        elif mod.type == "Analog Output":
            if rack.io_family == IO_FAMILY_FLEX5000:
                ladder = f"{mov}({tag},{rack.name}:{slot}:O.Ch{b:02d}.Data)"
            else:
                ladder = f"{mov}({tag},{rack.name}:{slot}:O.Ch{b}Data)"

        elif mod.type == "Thermocouple/RTD":
            ladder = f"{mov}({rack.name}:{slot}:I.Ch{b}Data,{tag})"

        elif mod.type == "Safety Input":
            if rack.io_family == IO_FAMILY_FLEX5000 or rack.io_family == IO_FAMILY_CLX:
                ladder = f"XIC({rack.name}:{slot}:I.Pt{b:02d}.Data)OTE({tag})"
            else:
                ladder = f"XIC({rack.name}:{slot}:I.Pt{b:02d}Data)OTE({tag})"
            
        elif mod.type == "Safety Output":
            if rack.io_family == IO_FAMILY_FLEX5000 or rack.io_family == IO_FAMILY_CLX:
                ladder = f"XIC({tag})OTE({rack.name}:{slot}:O.Pt{b:02d}.Data)"
            else:
                ladder = f"XIC({tag})OTE({rack.name}:{slot}:O.Pt{b:02d}Data)"
        else:
            raise ValueError(f"Unhandled module type {mod.type!r} in rack '{rack.name}' slot {mod.slot}.")

        rungs.append(_rung_xml(rung_num, "", ladder))
        rung_num += 1

        # Safety Input: extra status NOP rung
        if mod.type == "Safety Input":
            if rack.io_family == IO_FAMILY_FLEX5000 or rack.io_family == IO_FAMILY_CLX:
                status_ladder = f"XIC({rack.name}:{slot}:I.Pt{b:02d}.Status)NOP()"
            else:
                status_ladder = f"XIC({rack.name}:{slot}:I.Pt{b:02d}Status)NOP()"
            
            rungs.append(_rung_xml(rung_num, "", status_ladder))
            rung_num += 1

    return _routine_xml(mod.routine, rungs)


# ---------------------------------------------------------------------------
# Main chassis routine builder (always first in IO_Module_Status program)
# ---------------------------------------------------------------------------

def _build_main_chassis_routine(io_network_cards: list) -> str:
    rungs = []

    rung0_comment = (
        f"{_SEPARATOR}\n"
        f"This routine has an internal toggle bit to enable or disable the routines JSR instruction in the Subroutine Calls routine.\n"
        f"This bit is intended for commissioning and maintenance functions and must be high for normal line operations.\n"
        f"{_SEPARATOR}"
    )
    rungs.append(_rung_xml(0, rung0_comment, "XIC(JSR_ENABLE_A000_Main_Chassis)NOP()"))

    rung_num = 1
    for card in io_network_cards:
        card_comment = (
            f"{_SEPARATOR}\n"
            f"Main Chassis Slot {card.slot:02d} EN2T Module Status Logic\n"
            f"{_SEPARATOR}"
        )
        rungs.append(_rung_xml(rung_num, card_comment,
            f"GSV(Module,{card.name},EntryStatus,{card.name}._S_EntryStatus)"))
        rung_num += 1

        rungs.append(_rung_xml(rung_num, "", (
            f"XIO({card.name}._S_EntryStatus.12)"
            f"XIO({card.name}._S_EntryStatus.13)"
            f"XIC({card.name}._S_EntryStatus.14)"
            f"XIO({card.name}._S_EntryStatus.15)"
            f"OTE({card.name}._S_CommsOK)"
        )))
        rung_num += 1

        rungs.append(_rung_xml(rung_num, "", (
            f"[XIC(PLC._P_Module_Faults_Detect) XIO({card.name}._S_CommsOK) ,"
            f"XIC({card.name}._S_CommsFault) XIO(PLC._R_Module_Faults_Reset) ]"
            f"OTE({card.name}._S_CommsFault)"
        )))
        rung_num += 1

    return _routine_xml("A000_Main_Chassis", rungs)


# ---------------------------------------------------------------------------
# Module status routine builder (IO_Module_Status program)
# ---------------------------------------------------------------------------

def _mod_status_comment() -> str:
    return (
        f"{_SEPARATOR}\n"
        f"IO Module Status Subroutine\n"
        f"{_SEPARATOR}\n"
        f"This routine has an internal toggle bit to enable or disable the routines JSR instruction in the Subroutine Calls routine.\n"
        f"This bit is intended for commissioning and maintenance functions and must be high for normal line operations.\n"
        f"{_SEPARATOR}\n"
        f"Module Fault Detect Logic -- Master Control Relay\n"
        f"<<  When the MCR is disabled, the rung-condition-in is false for all the instructions inside this subroutine >>\n"
        f"{_SEPARATOR}"
    )


def _build_mod_status_routine(rack: Rack, io_card: str, software_version: str = "36.00") -> str:
    rungs = []
    rung_num = 0
    neq = _instr("NEQ", software_version)

    # MCR opening rung
    mcr_ladder = (
        f"XIC(JSR_ENABLE_{rack.name})"
        f"XIC(PLC._P_Module_Faults_Detect)"
        f"MCR()"
    )
    rungs.append(_rung_xml(rung_num, _mod_status_comment(), mcr_ladder))
    rung_num += 1

    # Communication module fault rung
    aent_comment = f"Module Fault Detect Logic\nCommunication Module"
    aent_ladder = (
        f"[XIO({io_card}._S_Fault) GSV(Module,{rack.name},FaultCode,{rack.name}._S_FaultCode) {neq}({rack.name}._S_FaultCode,0) ,\n"
        f"XIC({rack.name}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
        f"OTE({rack.name}._S_Fault)"
    )
    rungs.append(_rung_xml(rung_num, aent_comment, aent_ladder))
    rung_num += 1

    # Per-module fault rungs
    for mod in rack.modules:
        if not mod.routine:
            continue

        # Safety modules: fault logic lives in the safety program — add a placeholder NOP
        if mod.type in SAFETY_TYPES:
            comment = f"Module Fault Detect Logic is in Safety Program\nPlaceholder for Module {mod.slot}"
            rungs.append(_rung_xml(rung_num, comment, "NOP()"))
            rung_num += 1
            continue

        # Flex IO (1794): first module is slot 0; all others: first module is slot 1
        addr_slot = mod.slot - 1 if rack.io_family == IO_FAMILY_FLEX else mod.slot
        mod_comment = f"Module Fault Detect Logic\nModule {addr_slot}"

        if rack.io_family in (IO_FAMILY_FLEX, IO_FAMILY_CLX):
            if mod.type in DIGITAL_TYPES:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) XIC({rack.name}:I.SlotStatusBits.{addr_slot}) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
            else:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) GSV(Module,{mod.routine},FaultCode,{mod.routine}._S_FaultCode) {neq}({mod.routine}._S_FaultCode,0) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
        elif rack.io_family in (IO_FAMILY_FLEX5000):
            if mod.type in DIGITAL_TYPES:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) XIC({rack.name}:{addr_slot}:I.ConnectionFaulted) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
            else:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) GSV(Module,{mod.routine},FaultCode,{mod.routine}._S_FaultCode) {neq}({mod.routine}._S_FaultCode,0) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
        elif addr_slot < _POINT_IO_SLOTS_0_31_LIMIT:
            if mod.type in DIGITAL_TYPES:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) XIC({rack.name}:I.SlotStatusBits0_31.{addr_slot}) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
            else:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) GSV(Module,{mod.routine},FaultCode,{mod.routine}._S_FaultCode) {neq}({mod.routine}._S_FaultCode,0) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
        elif addr_slot < _POINT_IO_SLOTS_32_63_LIMIT:
            if mod.type in DIGITAL_TYPES:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) XIC({rack.name}:I.SlotStatusBits32_63.{addr_slot - 32}) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
            else:
                ladder = (
                    f"[XIO({rack.name}._S_Fault) GSV(Module,{mod.routine},FaultCode,{mod.routine}._S_FaultCode) {neq}({mod.routine}._S_FaultCode,0) ,\n"
                    f"XIC({mod.routine}._S_Fault) XIO(PLC._R_Module_Faults_Reset) ]\n"
                    f"OTE({mod.routine}._S_Fault)"
                )
        else:
            raise ValueError(
                f"Rack '{rack.name}' slot {addr_slot}: Point IO slot numbers >= {_POINT_IO_SLOTS_32_63_LIMIT} are not supported."
            )

        rungs.append(_rung_xml(rung_num, mod_comment, ladder))
        rung_num += 1

    return _routine_xml(rack.name, rungs)


# ---------------------------------------------------------------------------
# Safety module status routine builder (Safety_IO_Module_Status program)
# ---------------------------------------------------------------------------

def _build_safety_mod_status_routine(rack: Rack) -> str:
    rungs = []
    rung_num = 0

    comment = (
        f"{_SEPARATOR}\n"
        f"Safety IO Module Status Subroutine\n"
        f"{_SEPARATOR}\n"
    )

    mcr_ladder = f"XIC(SAFETY._Run_TMR.DN)MCR()"
    rungs.append(_rung_xml(rung_num, comment, mcr_ladder))
    rung_num += 1

    for mod in rack.modules:
        if not mod.routine or mod.type not in SAFETY_TYPES:
            continue
        ladder = (
            f"[XIC({rack.name}:{mod.slot}:I.ConnectionFaulted) ,"
            f"XIC({mod.routine}._S_Fault) XIO(SAFETY._R_Mod_Faults_Reset) ]"
            f"OTE({mod.routine}._S_Fault)"
        )
        rungs.append(_rung_xml(rung_num, "", ladder))
        rung_num += 1

    return _routine_xml(rack.name, rungs)


# ---------------------------------------------------------------------------
# Subroutine calls routine builder
# ---------------------------------------------------------------------------

def _build_calls_routine(routine_names: list, use_jsr_enable: bool = True) -> str:
    """Build the Subroutine_Calls routine with JSR rungs for each name."""
    rungs = [_rung_xml(0, "", "NOP()")]
    for i, name in enumerate(routine_names, start=1):
        if use_jsr_enable:
            rungs.append(_rung_xml(i, "", f"XIC(JSR_ENABLE_{name})JSR({name},0)"))
        else:
            rungs.append(_rung_xml(i, "", f"JSR({name},0)"))
    return _routine_xml("Subroutine_Calls", rungs)


# ---------------------------------------------------------------------------
# Local tag (JSR enable bit) builder
# ---------------------------------------------------------------------------

def _jsr_enable_tag(name: str, tag_class: str = "Standard") -> str:
    return _tag_xml(f"JSR_ENABLE_{name}", tag_class, "BOOL", "Local JSR Enable Bit", -1)


# ---------------------------------------------------------------------------
# Split-program bucket helper
# ---------------------------------------------------------------------------

_SPLIT_PROGRAM_NAMES = [
    "Digital_Input_Buffer",
    "Analog_Input_Buffer",
    "Digital_Output_Buffer",
    "Analog_Output_Buffer",
]


def _split_bucket(mod_type: str) -> str | None:
    """Map a module type to its split program name. Returns None for safety types."""
    if mod_type in OTHER_TYPES or mod_type == "Input":
        return "Digital_Input_Buffer"
    if mod_type in ("Analog Input", "Thermocouple/RTD"):
        return "Analog_Input_Buffer"
    if mod_type == "Output":
        return "Digital_Output_Buffer"
    if mod_type == "Analog Output":
        return "Analog_Output_Buffer"
    return None


# ---------------------------------------------------------------------------
# Top-level generate
# ---------------------------------------------------------------------------

def _build_standard_file(project: Project, target_name: str,
                          ctrl_tags: list, program_lines: list) -> list:
    """Build the lines for a standard-class L5X file containing one program."""
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<RSLogix5000Content SchemaRevision="1.0" SoftwareRevision="{project.software_version}" '
        f'TargetName="{target_name}" TargetType="Program" TargetClass="Standard" '
        f'ContainsContext="true" Owner="Windows User" '
        f'ExportOptions="References DecoratedData Context Dependencies ForceProtectedEncoding AllProjDocTrans">',
        f'<Controller Use="Context" Name="{project.controller_name}">',
        '<DataTypes Use="Context">',
        _UDT_QP_PLC_TAGS_v02,
        _UDT_QP_MODULE_TAGS_v02,
        '</DataTypes>',
        '<Tags Use="Context">',
        *[_tag_xml(card.name, "Standard", "QP_MODULE_TAGS_v02", f"{card.name}\nMain Enet Module", -1)
          for card in project.io_network_cards],
    ]
    lines.extend(ctrl_tags)
    lines.append('</Tags>')
    lines.append('<Programs Use="Context">')
    lines.extend(program_lines)
    lines.append('</Programs>')
    lines.append('</Controller>')
    lines.append('</RSLogix5000Content>')
    return lines


def _write_l5x(lines: list, path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write('\n'.join(lines))


def list_udts() -> list:
    """
    Parse the embedded UDT XML strings and return structured member data.

    Returns a list of dicts:
        {"name": str, "members": [{"name": str, "datatype": str, "bit": int|None, "description": str}]}
    Hidden backing members (ZZZZZZZZZZQP_*) are excluded.
    """
    import xml.etree.ElementTree as ET

    result = []
    for xml_str in _ALL_UDT_XML:
        root = ET.fromstring(xml_str)
        udt_name = root.get("Name")
        members = []
        for member in root.findall("Members/Member"):
            if member.get("Hidden") == "true":
                continue
            bit_raw = member.get("BitNumber")
            desc_el = member.find("Description")
            description = ""
            if desc_el is not None and desc_el.text:
                description = desc_el.text.strip()
            members.append({
                "name":        member.get("Name"),
                "datatype":    member.get("DataType"),
                "bit":         int(bit_raw) if bit_raw is not None else None,
                "description": description,
            })
        result.append({"name": udt_name, "members": members})
    return result


def generate(project: Project, output_dir: str, split_programs: bool = False) -> list:
    """
    Generate .l5x files into output_dir.
    Returns list of file paths written.

    Output files:
      IO_Files_Rev_<ts>.l5x             — IO_Buffer_Files (or 4 split programs) + IO_Module_Status
      Safety_IO_Files_Rev_<ts>.l5x      — Safety_IO_Buffer_Files + Safety_IO_Module_Status
                                          (only if safety modules present)

    When split_programs=True, IO_Buffer_Files is replaced by four programs:
      Digital_Input_Buffer, Analog_Input_Buffer, Digital_Output_Buffer, Analog_Output_Buffer
    """
    ts = datetime.now().strftime("%m%d%y_%H%M")
    written = []

    ctrl_tags, sfty_ctrl_tags = _build_ctrl_tags(project)

    buff_routine_names = []
    mod_status_names = ["A000_Main_Chassis"]
    sfty_routine_names = []
    sfty_mod_status_names = []

    buff_routines = []
    mod_routines = [_build_main_chassis_routine(project.io_network_cards)]
    sfty_routines = []
    sfty_mod_status_routines = []

    buff_local_tags = []
    mod_local_tags = [_jsr_enable_tag("A000_Main_Chassis")]
    sfty_local_tags = []
    sfty_mod_status_local_tags = []

    sfty_racks_with_modules = []  # list of (rack_name, [mod_routines]) for Modules context in safety file

    # Split-mode accumulators keyed by program name
    split_routine_names = {name: [] for name in _SPLIT_PROGRAM_NAMES}
    split_local_tags    = {name: [] for name in _SPLIT_PROGRAM_NAMES}
    split_routines      = {name: [] for name in _SPLIT_PROGRAM_NAMES}

    for rack in project.racks:
        mod_status_names.append(rack.name)
        mod_local_tags.append(_jsr_enable_tag(rack.name))
        mod_routines.append(_build_mod_status_routine(rack, rack.network_card, project.software_version))

        rack_sfty_mods = []
        for mod in rack.modules:
            if not mod.routine:
                continue

            if mod.type in SAFETY_TYPES:
                sfty_routine_names.append(mod.routine)
                sfty_local_tags.append(_jsr_enable_tag(mod.routine, "Safety"))
                sfty_routines.append(_build_buffer_routine(rack, mod, rack.network_card, project.software_version))
                rack_sfty_mods.append(mod.routine)
            elif split_programs:
                bucket = _split_bucket(mod.type)
                if bucket:
                    split_routine_names[bucket].append(mod.routine)
                    split_local_tags[bucket].append(_jsr_enable_tag(mod.routine))
                    split_routines[bucket].append(_build_buffer_routine(rack, mod, rack.network_card, project.software_version))
            else:
                buff_routine_names.append(mod.routine)
                buff_local_tags.append(_jsr_enable_tag(mod.routine))
                buff_routines.append(_build_buffer_routine(rack, mod, rack.network_card, project.software_version))

        if rack_sfty_mods:
            sfty_mod_status_names.append(rack.name)
            sfty_mod_status_local_tags.append(_jsr_enable_tag(rack.name, "Safety"))
            sfty_mod_status_routines.append(_build_safety_mod_status_routine(rack))
            sfty_racks_with_modules.append((rack.name, rack_sfty_mods))

    # ---- Buffer programs (combined or split) ----
    if split_programs:
        buffer_program_lines = []
        for prog_name in _SPLIT_PROGRAM_NAMES:
            buffer_program_lines += [
                f'<Program Use="Target" Name="{prog_name}" '
                'MainRoutineName="Subroutine_Calls" Class="Standard">',
                '<Tags>',
                *split_local_tags[prog_name],
                '</Tags>',
                '<Routines>',
                _build_calls_routine(split_routine_names[prog_name]),
                *split_routines[prog_name],
                '</Routines>',
                '</Program>',
            ]
        std_target_name = _SPLIT_PROGRAM_NAMES[0]
    else:
        buffer_program_lines = [
            '<Program Use="Target" Name="IO_Buffer_Files" '
            'MainRoutineName="Subroutine_Calls" Class="Standard">',
            '<Tags>',
            *buff_local_tags,
            '</Tags>',
            '<Routines>',
            _build_calls_routine(buff_routine_names),
            *buff_routines,
            '</Routines>',
            '</Program>',
        ]
        std_target_name = "IO_Buffer_Files"

    # ---- IO_Files (buffer programs + IO_Module_Status combined) ----
    combined_programs = [
        *buffer_program_lines,
        '<Program Use="Target" Name="IO_Module_Status" '
        'MainRoutineName="Subroutine_Calls" Class="Standard">',
        '<Tags>',
        *mod_local_tags,
        '</Tags>',
        '<Routines>',
        _build_calls_routine(mod_status_names),
        *mod_routines,
        '</Routines>',
        '</Program>',
    ]
    std_path = os.path.join(output_dir, f"IO_Files_Rev_{ts}.l5x")
    _write_l5x(_build_standard_file(project, std_target_name, ctrl_tags, combined_programs), std_path)
    written.append(std_path)

    # ---- Safety L5X (only if safety modules present) ----
    if sfty_ctrl_tags or sfty_routines:
        # Build Modules context entries: rack name first, then its module routine names
        modules_xml = []
        for rack_name, mod_routines_list in sfty_racks_with_modules:
            modules_xml.append(f'<Module Use="Reference" Name="{rack_name}">\n</Module>')
            for rname in mod_routines_list:
                modules_xml.append(f'<Module Use="Reference" Name="{rname}">\n</Module>')

        s_lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<RSLogix5000Content SchemaRevision="1.0" SoftwareRevision="{project.software_version}" '
            f'TargetName="Safety_IO_Module_Status" TargetType="Program" TargetClass="Safety" '
            f'ContainsContext="true" Owner="Windows User" '
            f'ExportOptions="References DecoratedData Context Dependencies ForceProtectedEncoding AllProjDocTrans">',
            f'<Controller Use="Context" Name="{project.controller_name}">',
            '<DataTypes Use="Context">',
            '</DataTypes>',
            '<Modules Use="Context">',
            *modules_xml,
            '</Modules>',
            '<Tags Use="Context">',
        ]
        s_lines.extend(sfty_ctrl_tags)
        s_lines.append('</Tags>')
        s_lines.append('<Programs Use="Context">')

        # Safety_IO_Module_Status program
        s_lines.append(
            '<Program Use="Target" Name="Safety_IO_Module_Status" TestEdits="false" '
            'MainRoutineName="Subroutine_Calls" Disabled="false" Class="Safety" UseAsFolder="false">'
        )
        s_lines.append('<Tags>')
        s_lines.extend(sfty_mod_status_local_tags)
        s_lines.append('</Tags>')
        s_lines.append('<Routines>')
        s_lines.append(_build_calls_routine(sfty_mod_status_names, use_jsr_enable=False))
        s_lines.extend(sfty_mod_status_routines)
        s_lines.append('</Routines>')
        s_lines.append('</Program>')

        # Safety_IO_Buffer_Files program
        s_lines.append(
            '<Program Use="Target" Name="Safety_IO_Buffer_Files" TestEdits="false" '
            'MainRoutineName="Subroutine_Calls" Disabled="false" Class="Safety" UseAsFolder="false">'
        )
        s_lines.append('<Tags>')
        s_lines.extend(sfty_local_tags)
        s_lines.append('</Tags>')
        s_lines.append('<Routines>')
        s_lines.append(_build_calls_routine(sfty_routine_names, use_jsr_enable=False))
        s_lines.extend(sfty_routines)
        s_lines.append('</Routines>')
        s_lines.append('</Program>')

        s_lines.append('</Programs>')
        s_lines.append('</Controller>')
        s_lines.append('</RSLogix5000Content>')

        sfty_path = os.path.join(output_dir, f"Safety_IO_Files_Rev_{ts}.l5x")
        _write_l5x(s_lines, sfty_path)
        written.append(sfty_path)

    return written
