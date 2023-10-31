import os
from dotenv import load_dotenv, find_dotenv
from dataclasses import dataclass

load_dotenv(find_dotenv())  # finds the .env file in the local config directory


@dataclass(frozen=True)  # makes the API keys immutable
class APIkeys:
    tenant_id: str = os.getenv('tenant_id')
    app_id: str = os.getenv('app_id')
    app_secret: str = os.getenv('app_secret')
