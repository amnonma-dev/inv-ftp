import xpo
import ulc
import a1
import time

def launch_xpo():
    print("preparing XPO ftp")
    xpo.main()
 
def launch_ulc():
    print("preparing ULC ftp")
    ulc.main()

def launch_a1():
    print("preparing A1 ftp")
    a1.main()

def quit():
    exit()

def errhandler(option_arg):    
    print("invalid option: "+str(option_arg))

actions = {
        1: launch_xpo,
        2: launch_ulc,
        3: launch_a1,
        4: quit
    }

while True:
    time.sleep(1)
    for key in sorted(actions):
        print (key, '=>', actions[key].__name__)
    selectedaction = input("please select an option from the list:    ")
    selected_action=None
    try:
        action_key=int(selectedaction)
        selected_action= actions.get(action_key)
    except:
        pass
    
    if  selected_action is None:
        errhandler(selectedaction)
        continue
    else:
        break
    
selected_action()
quit()
