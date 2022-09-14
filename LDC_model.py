


from system_simulator import SystemSimulator
from behavior_model_executor import BehaviorModelExecutor
from system_message import SysMessage
from definition import *

import pygsheets
from config import *

class LDCModel(BehaviorModelExecutor):
    def __init__(self, instance_time, destruct_time, name, engine_name):
        BehaviorModelExecutor.__init__(self, instance_time, destruct_time, name, engine_name)
        self.init_state("IDLE")
        self.insert_state("IDLE", Infinite)
        self.insert_state("WAKE", 2)
        self.insert_state("WAKE_", 2)
        self.insert_state("LABEL", 2)

        
  
        self.insert_input_port("msg")
        self.insert_input_port("sss")
        self.insert_input_port("lab")

        self.recv_msg = []
        
         
    def ext_trans(self,port, msg):
        print("msg check")
        if port == "msg":
            print("[Model] Received Msg")
            self._cur_state = "WAKE"
            self.recv_msg.append(msg.retrieve())
            self.cancel_rescheduling()

        if port == "lab":
            print("[Model] LABEL Msg")
            self._cur_state = "LABEL"
            self.recv_msg.append(msg.retrieve())
            self.cancel_rescheduling()

        if port == "sss":
            print("[Model] Received HWPNAME")
            self._cur_state = "WAKE_"
            self.recv_msg.append(msg.retrieve())
            self.cancel_rescheduling()
            

    def output(self):
        if self._cur_state == "WAKE":
            for _ in range(2000):
                if self.recv_msg:
                    print("check")
                    msg = self.recv_msg.pop(0)
                    
                    gc = pygsheets.authorize(service_file=GOOGLE_SERVIVE_KEY)
                    sh = gc.open(LAW_SHEETS_NAME)
                    wks = sh.worksheet('title','all_law')
                    law_df = wks.get_as_df()
                    try:
                        print("send")
                        wks.update_row((len(law_df)+2),msg,col_offset=0)
                    except:
                        print("pass")
                        pass
                
            #self.recv_msg = []
            pass
        if self._cur_state == "WAKE_":
            for _ in range(2000):
                if self.recv_msg:
                    print("hhhhhhh")
                    msg = self.recv_msg.pop(0)
                    print(msg)
                    gc = pygsheets.authorize(service_file=GOOGLE_SERVIVE_KEY)
                    sh = gc.open(LAW_SHEETS_NAME)
                    wks = sh.worksheet('title','한글파일_종류_정리')
                    df=wks.get_as_df()
                    wks.update_value('A' + str(len(df)+2), msg[-1])
        
        if self._cur_state == "LABEL":
            for _ in range(3000):
                if self.recv_msg:
                    print("laberrr")
                    msg = self.recv_msg.pop(0)
                    print("____")
                    print(msg)
                    gc = pygsheets.authorize(service_file=GOOGLE_SERVIVE_KEY)
                    sh = gc.open(LAW_SHEETS_NAME)
                    wks = sh.worksheet('title','all_law')
                    wks.update_value('C'+msg[0],msg[1])
                    wks.update_value('D'+msg[0],msg[2])
                    wks.update_value('E'+msg[0],msg[3])


    def int_trans(self):
        if self._cur_state == "WAKE":
            self._cur_state = "IDLE"
    
    def __del__(self):
        pass