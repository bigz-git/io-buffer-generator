from dataclasses import dataclass, field


DIGITAL_TYPES = {"Input", "Output"}
ANALOG_TYPES = {"Analog Input", "Analog Output", "Thermocouple/RTD"}
SAFETY_TYPES = {"Safety Input", "Safety Output"}
OTHER_TYPES = {"Other"}
ALL_MODULE_TYPES = sorted(DIGITAL_TYPES | ANALOG_TYPES | SAFETY_TYPES | OTHER_TYPES)

MODULE_TYPE_DROPDOWN = "Input,Output,Safety Input,Safety Output,Analog Input,Analog Output,Thermocouple/RTD,Other"

IO_FAMILY_POINT = "1734"   # Point IO
IO_FAMILY_FLEX  = "1794"   # Flex IO
IO_FAMILY_CLX   = "1756"   # ControlLogix IO
IO_FAMILY_DROPDOWN = f"{IO_FAMILY_POINT},{IO_FAMILY_FLEX},{IO_FAMILY_CLX}"


@dataclass
class Bit:
    index: int        # 0-indexed bit/channel within module
    tag: str          # buffer tag name, e.g. "CONV_01_IN.0" or "CONV_01_AIN[0]"
    description: str  # tag description
    drawing: str      # drawing file name (stored on first bit row of each module)


@dataclass
class Module:
    slot: int             # slot number as entered (1-indexed)
    type: str             # "Input", "Output", "Safety Input", etc.
    routine: str          # PLC routine name
    bits: list = field(default_factory=list)  # list[Bit]


@dataclass
class Rack:
    name: str
    io_family: str = IO_FAMILY_POINT  # "1734" (Point IO) or "1794" (Flex IO)
    modules: list = field(default_factory=list)  # list[Module]


@dataclass
class Project:
    software_version: str
    controller_name: str
    io_network_card: str
    project_number: str = ""
    project_description: str = ""
    racks: list = field(default_factory=list)  # list[Rack]
