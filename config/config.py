import os
from pathlib import Path
from dotenv import load_dotenv, dotenv_values
import logging
import sys

dotenv_path = os.path.join(Path(__file__).parent.parent, ".env")

if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)

config_env = os.environ if not os.path.exists(dotenv_path) else dotenv_values(".env")

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(logging.StreamHandler(sys.stdout))
logger.info("Screen shots is starting up")

