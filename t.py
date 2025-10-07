import pandas as pd
import openai
from pydantic import BaseModel
import json, os
from dotenv import load_dotenv

#Load environment variables
load_dotenv()
openai_api_key = os.getenv("openai_api_key")

#Initialize OpenAI client
client = openai.OpenAI(api_key=openai_api_key)

#Pydantic model for structured output
class EmailExtraction(BaseModel):
    email: str
    name: str
    company: str

#Excel to CSV conversion
excel_path = r"C:\Users\Sri Nithi Murali\PycharmProjects\CobuildX.ai1\emails.xlsx"
csv_path = excel_path.replace(".xlsx", ".csv")
df = pd.read_excel(excel_path)
df.to_csv(csv_path, index=False)
print(f"Converted Excel to CSV: {csv_path}")

#List to store all parsed results
all_parsed = []

#Loop through emails and parse
for email in df['Email']:
    email = str(email).strip()
    if not email:
        print("Empty email, skipping row")
        continue

    system_msg = f"""
    Extract the full name and company from this email in JSON format with keys:
    {{
      "email": "...",
      "name": "...",
      "company": "..."
    }}

    Rules for extraction:
    1. The 'email' field should be exactly as given.
    2. The 'name' field should only contain alphabetic characters (A-Z, a-z). 
       - If the local part of the email contains numbers mixed with letters (e.g., user9876), discard the numbers and return only the letters as the name.
       - If the local part contains only numbers or only special characters, set the name as "NA".
       - If the name can’t be inferred from the email, set it as "NA".
    3. The 'company' field should be inferred from the domain name (e.g., "infosys.com" -> "Infosys").
       - If the domain name is not recognizable, use "NA".
    4. Always return strictly valid JSON with the keys: "email", "name", "company".
    5. Do not include extra explanations, markdown formatting, or backticks. Only return JSON.
    6. Be robust: handle emails with unusual patterns, extra dots, underscores, numbers, or symbols.
    7. Examples:
       - "john.doe@wipro.com" -> {{"email":"john.doe@wipro.com","name":"John Doe","company":"Wipro"}}
       - "user9876@weirdname.biz" -> {{"email":"user9876@weirdname.biz","name":"User","company":"Weirdname"}}
       - "12345@unknown.xyz" -> {{"email":"12345@unknown.xyz","name":"NA","company":"Unknown"}}
       - "_._@strange-domain.org" -> {{"email":"_._@strange-domain.org","name":"NA","company":"Strangedomain"}}

    Email to parse: {email}
    """

    try:
        #Use structured output via nightly .parse()
        response = client.responses.parse(
            model="gpt-4o-mini-2024-07-18",
            input=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": system_msg},
            ],
            text_format=EmailExtraction,
            temperature=0.0  # Keep low for consistent extraction
        )

        #Access parsed Pydantic object
        parsed: EmailExtraction = response.output_parsed  # Correct attribute
        parsed_dict = parsed.model_dump()
        all_parsed.append(parsed_dict)

        print(f"{parsed_dict}")

    except Exception as e:
        print(f"Error for {email}: {e}")

#Write all parsed results to JSON
output_file = "parsed_emails.json"
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(all_parsed, f, indent=4, ensure_ascii=False)

print(f"All parsed data saved to {output_file}")
