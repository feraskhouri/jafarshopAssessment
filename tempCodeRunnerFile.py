import os
import re
import time
import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI

# Load environment variables from .env
load_dotenv()
client = OpenAI()

def get_weight_from_gpt(product_name):
    prompt = f"What is the weight of the product '{product_name}' in grams? Just give the number followed by 'g'."

    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",  # free-tier friendly model
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        reply = response.choices[0].message.content.strip()
        print(f"GPT response: {reply}")

        match = re.search(r"(\d+\.?\d*)\s*g", reply.lower())
        if match:
            return float(match.group(1)), "API"
        else:
            return 0, "API (Not Found)"
    except Exception as e:
        print(f"❌ Error for '{product_name}': {e}")
        return 0, "API (Error)"

def main():
    input_file = "product_detailsS_updated.xlsx"
    output_file = "product_weights_with_api.xlsx"

    df = pd.read_excel(input_file)

    for index, row in df[df['Weight'] == 0].iterrows():
        product_name = row["Product Name (EN)"]
        print(f"🔍 Querying: {product_name}")

        weight, method = get_weight_from_gpt(product_name)
        df.loc[index, 'Weight'] = weight
        df.loc[index, 'Detection Method'] = method

        time.sleep(1.5)  # avoid hitting rate limits

    df.to_excel(output_file, index=False)
    print(f"✅ Done! Updated file saved as '{output_file}'")

if __name__ == "__main__":
    main()
