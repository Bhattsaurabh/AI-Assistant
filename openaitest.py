import os
import openai
from config import apikey

openai.api_key = apikey

response = openai.Completion.create(
  model="text-davinci-003",
  prompt=prompt,
  temperature=0.7,
  max_tokens=256,
  top_p=1,
  frequency_penalty=0,
  presence_penalty=0
)
print(response)

"""
{
  "id": "chatcmpl-8EBiHJE3CPAe1uxH7VnHJpZyYBkgc",
  "object": "chat.completion",
  "created": 1698392309,
  "model": "gpt-3.5-turbo-0613",
  "choices": [
    {
      "index": 0,
      "message": {
        "role": "assistant",
        "content": "am confident in the capabilities of the team, and I have full faith that they will continue to excel in their roles and contribute to the company's success.\n\nPlease let me know if there are any formalities or procedures I need to follow regarding my resignation. I am open to discussing a suitable handover plan and wrapping up any outstanding tasks.\n\nI want to extend my sincere gratitude to you for your guidance, mentorship, and support throughout my time here at [Company Name]. It has been an honor to work with such a dedicated and professional team. I am grateful for the opportunities and experiences that have shaped me both personally and professionally.\n\nThank you for your understanding. I wish you and the entire team at [Company Name] continued success in all future endeavors.\n\nKind regards,\n\n[Your Name]"
      },
      "finish_reason": "stop"
    }
  ],
  "usage": {
    "prompt_tokens": 280,
    "completion_tokens": 159,
    "total_tokens": 439
  }
}
"""