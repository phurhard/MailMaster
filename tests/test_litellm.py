from litellm import completion
from dotenv import load_dotenv
import os

load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

prompt = "Hello, respond with just the word 'Hi'."

try:
    response = completion(
        model=os.getenv("OPENAI_MODEL"),
        messages=[{"role": "user", "content": prompt}],
        api_key=api_key
    )
    print("Response object type:", type(response))
    print("Response object content:", response)
    
    # Try accessing message content using object notation (Pydantic style)
    try:
        content_obj = response.choices[0].message.content
        print("Object notation success! Content:", content_obj)
    except Exception as e1:
        print(f"Object notation failed: {e1}")
        
        # Try accessing message content using dictionary notation
        try:
            content_dict = response["choices"][0]["message"]["content"]
            print("Dictionary notation success! Content:", content_dict)
        except Exception as e2:
            print(f"Dictionary notation failed: {e2}")

except Exception as main_e:
    print(f"Main completion failed: {main_e}")
