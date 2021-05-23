import write_to_sheet as ws
import schedule
import time
import datetime

def main():
    run = ws.Run()
    while True:
        i=0
        x = datetime.datetime.now()
        hr = x.strftime("%H")
        mini = x.strftime("%M")
        if hr == "10" and mini == "22":
            print("running job ...")
            run.job(i)      
            time.sleep(60)
        else:
            pass
        






if __name__=="__main__":
    main()