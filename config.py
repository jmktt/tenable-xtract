import os
from dotenv import load_dotenv

def carregar_env():
    load_dotenv()
load_dotenv()

CLIENTES = {
    'client1': {
        'access_key': os.getenv('client1_ACCESS_KEY'),
        'secret_key': os.getenv('client1_SECRET_KEY')
    },
    'client2': {
        'access_key': os.getenv('client2_ACCESS_KEY'),
        'secret_key': os.getenv('client2_SECRET_KEY')
    },
    'client3': {
        'access_key': os.getenv('client3_ACCESS_KEY'),
        'secret_key': os.getenv('client3_SECRET_KEY')
    }
}

STATUS_ONLINE = 'on'
STATUS_OFFLINE = 'off'
TABLE_STYLE = 'Table Style Dark 1'  # DARKMODE

PLUGINS_SOFTWARE_INVENTORY = {
    20811: "Windows Software",
    22869: "Linux Packages"
}

EXPLOITDB_EXPLOITS_CSV_URL = "https://gitlab.com/exploit-database/exploitdb/-/raw/main/files_exploits.csv"
EXPLOITDB_SHELLCODES_CSV_URL = "https://gitlab.com/exploit-database/exploitdb/-/raw/main/files_shellcodes.csv"

EXPLOITDB_EXPLOITS_LOCAL_CSV_PATH = os.path.join(os.getcwd(), 'exploitdb_exploits.csv')
EXPLOITDB_SHELLCODES_LOCAL_CSV_PATH = os.path.join(os.getcwd(), 'exploitdb_shellcodes.csv')

DELETE_EXPLOITDB_CSV_AFTER_USE = True

# Banners
banner = '''

               ▐▄• ▄ ▄▄▄▄▄▄▄▄   ▄▄▄·  ▄▄· ▄▄▄▄▄
                █▌█▌▪•██  ▀▄ █·▐█ ▀█ ▐█ ▌▪•██  
                ·██·  ▐█.▪▐▀▀▄ ▄█▀▀█ ██ ▄▄ ▐█.▪
               ▪▐█·█▌ ▐█▌·▐█•█▌▐█ ▪▐▌▐███▌ ▐█▌·
               •▀▀ ▀▀ ▀▀▀ .▀  ▀ ▀  ▀ ·▀▀▀  ▀▀▀ 
                                                                           
'''

banner1 = '''
      
      $$\   $$\ $$$$$$$$\ $$$$$$$\   $$$$$$\   $$$$$$\ $$$$$$$$\ 
      $$ |  $$ |\__$$  __|$$  __$$\ $$  __$$\ $$  __$$\\__$$  __|
      \$$\ $$  |   $$ |   $$ |  $$ |$$ /  $$ |$$ /  \__|  $$ |   
       \$$$$  /    $$ |   $$$$$$$  |$$$$$$$$ |$$ |        $$ |   
       $$  $$<     $$ |   $$  __$$< $$  __$$ |$$ |        $$ |   
      $$  /\$$\    $$ |   $$ |  $$ |$$ |  $$ |$$ |  $$\   $$ |   
      $$ /  $$ |   $$ |   $$ |  $$ |$$ |  $$ |\$$$$$$  |  $$ |   
      \__|  \__|   \__|   \__|  \__|\__|  \__| \______/   \__|   
                                                                                                                                                                                                                                     
'''

banner2 = '''\033[1;30m
      
      ▒██   ██▒▄▄▄█████▓ ██▀███   ▄▄▄       ▄████▄  ▄▄▄█████▓
      ▒▒ █ █ ▒░▓  ██▒ ▓▒▓██ ▒ ██▒▒████▄    ▒██▀ ▀█  ▓  ██▒ ▓▒
      ░░  █   ░▒ ▓██░ ▒░▓██ ░▄█ ▒▒██  ▀█▄  ▒▓█    ▄ ▒▓█    ▄ ▒ ▓██░ ▒░
       ░ █ █ ▒ ░ ▓██▓ ░ ▒██▀▀█▄  ░██▄▄▄▄██ ▒▓▓▄ ▄██▒░ ▓██▓ ░ 
      ▒██▒ ▒██▒  ▒██▒ ░ ░██▓ ▒██▒ ▓█   ▓██▒▒ ▓███▀ ░  ▒██▒ ░ 
      ▒▒ ░ ░▓ ░  ▒ ░░   ░ ▒▓ ░▒▓░ ▒▒   ▓▒█░░ ░▒ ▒  ░  ▒ ░░   
      ░░   ░▒ ░    ░      ░▒ ░ ▒░  ▒   ▒▒ ░  ░  ▒       ░    
       ░    ░    ░        ░░   ░   ░   ▒   ░          ░      
       ░    ░              ░           ░  ░░ ░               
                                           ░                 \033[0m
'''

banner3= '''\033[1;37m\033[5m
      ,-.
     / \  `.  __..-,O
    :   \ --''_..-'.'
    |    . .-' `. '.
    :     .     .`.'
     \     `.  /  ..
      \      `.   ' .
       `,       `.   I
      ,|,`.        `-.I
     '.||  ``-...__..-`
      |  |
      |__|
      /||I
     //||\\
    // || \\
 __//__||__\\__ \033[0m\033[0;31mXTRACT\033[0m
'''

banners = [banner, banner1, banner2, banner3]