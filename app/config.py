from dotenv import load_dotenv
load_dotenv()

import os

SECRET_KEY = os.getenv("SECRET_KEY")
ACCESS_TOKEN_EXPIRE_MINUTES = int(os.getenv("ACCESS_TOKEN_EXPIRE_MINUTES", 30))
LLM_API_KEY = os.getenv("LLM_API_KEY", "")
