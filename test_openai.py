from openai import OpenAI

client = OpenAI(
  api_key="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"
)

completion = client.chat.completions.create(
  model="gpt-4o-mini",
  store=True,
  messages=[
    {"role": "user", "content": "write a haiku about ai"}
  ]
)

print(completion.choices[0].message);
