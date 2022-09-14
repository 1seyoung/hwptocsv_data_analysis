from config import *


from LD_mgr import LDManager


from system_simulator import SystemSimulator
from behavior_model_executor import BehaviorModelExecutor
from system_message import SysMessage


from definition import *

import os

# System Simulator Initialization
se = SystemSimulator()

se.register_engine("sname", SIMULATION_MODE, TIME_DENSITY)

se.get_engine("sname").insert_input_port("start")
se.get_engine("sname").insert_input_port("msg")
se.get_engine("sname").insert_input_port("lab")
se.get_engine("sname").insert_input_port("sss")



# Telegram Manager Initialization
ld = LDManager(se.get_engine("sname"))


# Monitoring System Start
ld.start()