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
        'secret_key': os.getenv('AUSTRAL_SECRET_KEY')
    },
    'client3': {
        'access_key': os.getenv('client3_ACCESS_KEY'),
        'secret_key': os.getenv('client3_SECRET_KEY')
    }
}
