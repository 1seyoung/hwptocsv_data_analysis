from lib2to3.pygram import python_grammar_no_print_statement
from operator import index
from config import *

from system_simulator import SystemSimulator
from behavior_model_executor import BehaviorModelExecutor
from system_message import SysMessage
from definition import *

from LDC_model import LDCModel as LDC

import pygsheets
import signal
import glob
from pathlib import Path

import struct
import zlib
from pprint import pprint
import olefile
import re
import sys,os
import pandas as pd

'''
# Simulation Configuration
se.get_engine("sname").simulate()'''

class LDManager():   
    def __init__(self,engine):
        self.se = engine
		

        signal.signal(signal.SIGINT,  self.signal_handler)
        signal.signal(signal.SIGABRT, self.signal_handler)
        signal.signal(signal.SIGTERM, self.signal_handler)

        self.is_terminating = False

        model = LDC(0, Infinite, "ld", "sname")


        self.se.coupling_relation(None, "start", model, "start")
        self.se.coupling_relation(None, "msg", model, "msg")
        self.se.coupling_relation(None, "sss", model, "sss")
        self.se.coupling_relation(None, "lab", model, "lab")

        self.se.register_entity(model)
        self.a =0

    def dir_list_csv(self):
        #모든 파일 리스트 출력 +path
        if self.a ==1:

            path_list = []
            filename_list = []
            #format=['.hwp','.pdf','.XLS','.xlsx','.zip','.dwg','.cbt','.xlsm','.xls'] -> all_file
            format=['.hwp'] #-> hwp
            df = pd.DataFrame(columns=['path', 'filename'])
            for (path, dir, files) in os.walk(DATA_PATH):
                for filename in files:
                    ext = os.path.splitext(filename)[-1]
                    if ext in format:
                        print("%s/%s" % (path, filename))
                        df=df.append({'path' : path , 'filename' : filename} , ignore_index=True)

            df.to_csv(CSVFILE_PATH, index=False)
        else:
            pass

    def law_in_hwp_csv(self,foldername,route):


        if self.a ==1:

            
        
            engineering_data = route
            hwps = glob.glob(engineering_data + '/*.hwp')
            hwp_line=[]
            law_data_df = pd.DataFrame(columns=['hwp_name','law_name','count'])
            for hwp in hwps:
                
                line ={}
                line_count =[]
                try:
                    print(hwp)

                    f = olefile.OleFileIO(hwp)

                    dirs = f.listdir()
                    
                    # 문서 압축 여부 확인
                    header = f.openstream('FileHeader')
                    header_data = header.read()
                    
                    is_compressed = (header_data[36] & 1) == 1
                    # Body Sections 불러오기
                    # nums: 'Section0'이라고 적혀 있을 때 '0'만 수집하는 공간
                    nums = []
                    for d in dirs:
                        if d[0] == 'BodyText':
                            nums.append(int(d[1][len("Section"):]))
                    sections = ["BodyText/Section" + str(x) for x in sorted(nums)]
                    
                    # 전체 텍스트 추출
                    text = ''
                    for section in sections:
                        # openstream('PrvText')는 미리보기 였음...
                        # 모든 텍스트를 다 뽑으려면 BodyText 안의 Sections에서 뽑아야 했음
                        # section ==> ex) "BodyText/Section0", "BodyText/Section1", ...
                        
                        bodytext = f.openstream(section)
                        data = bodytext.read()
                        
                        if is_compressed:
                            unpacked_data = zlib.decompress(data, -15)
                        else:
                            unpacked_data = data
                        # 각 Section에서 text 추출
                        section_text = ''
                        i = 0 
                        size = len(unpacked_data)
                        while i < size:
                            header = struct.unpack_from('<I', unpacked_data, i)[0]
                            rec_type = header & 0x3ff
                            rec_len = (header >> 20) & 0xfff
                            if rec_type in [67]:
                                rec_data = unpacked_data[i+4:i+4+rec_len]
                                section_text += rec_data.decode('utf-16')
                                section_text += '\n'
                            i += 4 + rec_len
                        text += section_text
                        text += '\n'
                        
                    # start_end1 ==> 꺽새로 이루어진 법 추출
                    # start_end2 ==> 대괄호로 이루어진 법 추출
                    start_end1 = [None, None]
                    start_end2 = [None, None]
                    #start_end3 = [None, None]
                    for i in range(len(text)):
                        #print(text)
                        if ord(text[i]) == 12300:
                            start_end1[0] = i
                        elif ord(text[i]) == 12301:
                            if start_end1[0] == None:
                                continue
                            start_end1[1] = i
                            for word in ['법', '규칙', '규정', '규약', '요령', '기준', '근거']:
                                if word in text[start_end1[0]:start_end1[1]+1]:
                                    print(text[start_end1[0]+1: start_end1[1]])
                                    law_name = text[start_end1[0]+1: start_end1[1]]
                                    line_count.append(law_name)
                                    line[law_name] =0
                                    
                                    
                                    break
                            start_end1 = [None, None]
                        elif text[i] == '[':
                            start_end2[0] = i
                        elif text[i] == ']':
                            start_end2[1] = i
                            for word in ['법', '규칙', '규정', '규약', '요령', '기준', '근거']:
                                if word in text[start_end2[0]: start_end2[1]+1]:
                                    print(text[start_end2[0]+1: start_end2[1]])
                                    law_name = text[start_end2[0]+1: start_end2[1]]
                                    line_count.append(law_name)
                                    line[law_name] =0
                                    break


                    hwp_ = hwp.split('/')
                    hwp_name = hwp_[-1]
                    #count 
                    for law in line_count:
                        try: line[law] +=1
                        except: line[law] =1
                    #df 만들기
                    for t,s in zip(line.keys(),line.values()):
                        
                        msg = []

                        preprocessing_hwp = re.sub('[;:\*?!~`’^\-_+<>@\#$%&=#/}※ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅎㅊㅋㅌㅍㅠㅜ]','', hwp_name)
                        preprocessing_hwp = re.sub('.hwp','',preprocessing_hwp)
                        preprocessing_hwp = re.sub('ok','',preprocessing_hwp)
                        msg.append(foldername)
                        msg.append(preprocessing_hwp)
                        msg.append(t)
                        msg.append(s)
                        print("sddddd")
                        self.se.insert_custom_external_event("msg", msg)                    
                except:
                    pass
        else:
            pass            
                


        print('-'*70)
        print()
        




    def start(self) -> None:
        print("start")

        if '--data' in sys.argv:
            self.a =1
            self.dir_list_csv()

            raw_data = pd.read_csv(CSVFILE_PATH, engine='python')
            path_set = set(raw_data['path'].tolist())
            path_list= list(path_set)

            for route in path_list:
                
                file_path = route
                
                

                self.law_in_hwp_csv(file_path ,route)
            self.se.simulate()

        if '--cleanup' in sys.argv:
            self.a =0
            gc = pygsheets.authorize(service_file=GOOGLE_SERVIVE_KEY)
            sh = gc.open(LAW_SHEETS_NAME)
            wks = sh.worksheet('title','all_law')
            law_df = wks.get_as_df()
            msg=[]
            hwplist=list(set(law_df['hwp_name'].tolist()))
            for hwp in hwplist:
                print(hwp)
                msg.append(hwp)
                self.se.insert_custom_external_event("sss", msg)
            self.se.simulate()
        
        if '--labeling' in sys.argv:
            print("label check")
            
            self.a = 0
            gc = pygsheets.authorize(service_file=GOOGLE_SERVIVE_KEY)
            sh = gc.open(LAW_SHEETS_NAME)
            wks = sh.worksheet('title','한글파일_종류_정리')
            wkss=sh.worksheet('title','all_law')
            label_df =wks.get_as_df()
            law_df = wkss.get_as_df()

            for i in range (len(label_df)):
                a = (label_df.loc[[i],['hwp','LEVEL2','LEVEL3','LEVEL4']]).values.tolist()
                a=a[0]
                hwpdata = law_df.index[law_df['hwp_name']==a[0]].tolist()

                for i in range(len(hwpdata)):

                    msg=[]
                    idx=hwpdata[i]+2
                    idx=str(idx)
                    
                    msg.append(idx)
                    msg.append(a[1])
                    msg.append(a[2])
                    msg.append(a[3])
                    self.se.insert_custom_external_event("lab", msg)


                    
                    

            self.se.simulate()



	

		

    def signal_handler(self, sig, frame):
        print("Terminating Monitoring System")
		
        if not self.is_terminating:
            self.is_terminating = True
            self.updater.stop()
            del self.se
		
        sys.exit(0)




