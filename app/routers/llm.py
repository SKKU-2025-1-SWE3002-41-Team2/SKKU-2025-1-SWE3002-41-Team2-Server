
import os
from dotenv import load_dotenv
import openai


loaddotenv()

async def process_natural_language_command(command: str) -> str:
    api_key = os.getenv("LLM_API_KEY")
    if not api_key:
        raise ValueError("LLM_API_KEY is not set in environment variables.")

    openai.api_key = api_key

    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are an assistant that helps with Excel tasks."},
            {"role": "user", "content": command}
        ],
        temperature=0.5,
        max_tokens=1000
    )

    result = response['choices'][0]['message']['content'].strip()
    return result
