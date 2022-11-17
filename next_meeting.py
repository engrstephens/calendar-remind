
import time
from datetime import datetime
from datetime import timedelta
import toaster
import traceback
import random

import win32com.client
#https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.meetingitem?view=outlook-pia



def speak(speaker,msg):
    print("speaker",speaker)
    print("msg",msg)
    vcs = speaker.GetVoices()
    print("vcs",vcs)
    print("vcs",len(vcs))
    voice=random.randrange(1,len(vcs))
    print("voice",voice)
    #print("voice",vcs.Item(voice))
    #speaker.SetVoice(vcs.Item(voice))
    speaker.Speak(msg)

def get_calendar(outlook,begin,end, foldername=None):
        
    if foldername == None:
        calendar = outlook.getDefaultFolder(9).Items
    else:
        calendar = outlook.getDefaultFolder(9).Folders[foldername].Items
    
    #restriction = "[Sensitivity]<>'Private' AND [Start] >= '" + begin.strftime('%m/%d/%Y %H:%M') + "' AND [END] <= '" + end.strftime('%m/%d/%Y %H:%M') + "'"
    restriction = "[Sensitivity]<>'Private' AND [Start] >= '" + begin.strftime('%m/%d/%Y %H:%M') + "' AND [END] <= '" + end.strftime('%m/%d/%Y %H:%M') + "'"
    #print(restriction)
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    calendar = calendar.Restrict(restriction) 
    return calendar


def main():
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    last_time=datetime.min
    c=[]
    old_c=None
    t = toaster.ToastNotifier()
    while True:
        try:
            #Get the calendar
            now=datetime.now()
            time_since_update=now-last_time
            if time_since_update.total_seconds()> 5*60:
                print("fetching next")
                last_time=now
                c=get_calendar(outlook, now- timedelta(minutes=1), now+ timedelta(hours=8))
            if c.GetFirst()!=None:
                a=c[0] #First appointment
                    
                #Grab the COM datetime
                start=datetime.strptime(str(a.Start),"%Y-%m-%d %H:%M:%S%z")
                start = start.replace(tzinfo=None)

                #Figure out how much time is left until the appointmenmt
                delta=start-now
                seconds=delta.total_seconds()
                minutes=seconds/60.0
                minutes=int(minutes)
                seconds=int(seconds)
                subject=str(a.Subject)

                #Print debugging info
                print(subject,minutes,seconds)

                #If we haven't ran yet or there is a new meeting, announce it
                if old_c == None:
                    init=True
                elif old_c.EntryID!=a.EntryID:
                    init=True
                else:
                    init=False
                
                if minutes>60*10:
                    time.sleep(60)
                    continue
                
                #If we are init or the minutes is one of the key times I want an announcement
                if minutes>=2 and (((minutes % 30) == 0 or minutes in [45,20,15,10,5,3,2]) or (init==True)):
                    msg="T -"+ str(minutes)+ " minutes"
                    speak(speaker,"T -"+ str(minutes)+ " minutes")            
                    t.show_toast(subject,msg,duration=60)
                    #The duration is set here for 60 seconds so that it sleeps for 60 seconds
                    time.sleep(45)
                #If we are init or the seconds is one of the key times I want an announcement
                elif  minutes <2 and (seconds in [60,30,20,10] or init==True):
                        msg="T -"+ str(seconds)+ " seconds"
                        speak(speaker,msg)  #Just print seconds since we are counting down!
                        t.show_toast(subject,msg,duration=1)
                        time.sleep(0.5)
                elif seconds<5 and seconds>0:
                    speak(speaker,"T - 5")
                    speak(speaker,"4")
                    speak(speaker,"3")
                    speak(speaker,"2")
                    speak(speaker,"1")
                    time.sleep(5)
                else:
                    time.sleep(1)
                
                old_c=a
            else:
                time.sleep(1)
        except Exception:
            print(traceback.format_exc())
            old_c=None
            speak(speaker,"oops")
            time.sleep(1)
            return

while True:
    main()