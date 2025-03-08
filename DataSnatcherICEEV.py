import os
import json
import requests
import re
import statistics
import sys
import random

# Load API key from environment variable
API_KEY = os.getenv("XAI_API_KEY", "xai-xYIlIJrJlOzrENGLpKEMtNBkZVTjeIu4bRSLp2HP0CgnSHaPuppQ2ywNBrGAfDEIdcfJpOUKoNJYysPc")
API_URL = "https://api.x.ai/v1/chat/completions"
MODEL = "grok-2-1212"

# Function to query the API
def query_x_api(prompt):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": prompt}]
    }
    
    response = requests.post(API_URL, headers=headers, json=payload)
    
    print(f"\n📡 API Request: {prompt}")
    print(f"🔄 Response Status: {response.status_code}")
    
    if response.status_code == 200:
        try:
            response_json = response.json()
            print(f"📩 Raw API Response:\n{json.dumps(response_json, indent=4)}")
            return response_json
        except json.JSONDecodeError:
            print("⚠️ Error: Response is not valid JSON.")
            print("Raw response:", response.text)
            return None
    else:
        print("❌ API Error:", response.status_code, response.text)
        return None

# Function to extract a single numeric value from content
def extract_value(content):
    content = content.strip()
    # Strip everything except digits and decimal points
    numeric_str = re.sub(r'[^\d.]', '', content)
    if not numeric_str:
        print(f"⚠️ No numeric data in '{content}'. Defaulting to 0.")
        return 0
    try:
        # Convert to float first to handle decimals
        value = float(numeric_str)
        # If the value is unusually large (e.g., > 1000), assume it's in dollars and convert to thousands
        if value > 1000:  # Likely in dollars, convert to thousands
            value = value / 1000
            print(f"⚠️ Value {numeric_str} seems to be in dollars, converting to thousands: {value}")
        # Return as int if no decimal, otherwise float
        return int(value) if value.is_integer() else value
    except ValueError:
        print(f"⚠️ Could not parse '{content}' as a number. Defaulting to 0.")
        return 0

# Extract content safely (expecting single values)
def extract_content(response):
    if response and "choices" in response and len(response["choices"]) > 0:
        content = response["choices"][0]["message"]["content"].strip()
        return extract_value(content)
    print("⚠️ No valid content found in response.")
    return 0

# Step 1: Read the existing car_kpi_data.json or fetch new data if needed
input_filename = "car_kpi_data.json"
data = {}  # Initialize data dictionary to avoid undefined variable issues
try:
    with open(input_filename, "r") as json_file:
        data = json.load(json_file)
    kpi_list = data["top_5_KPIs"]
    # Check for the updated key name and fallback to old key if needed
    if "top_cars_US_2024" in data:
        car_list = data["top_cars_US_2024"]
    elif "top_bought_cars_US" in data:  # For backward compatibility
        car_list = data["top_bought_cars_US"]
    else:
        raise KeyError("No valid car list key found")
    if len(car_list) != 20:  # Now expecting 20 cars (10 ICE + 10 EV)
        raise KeyError("Need 20 cars (10 ICE + 10 EV)")
    print(f"✅ Loaded data from {input_filename}")
    print(f"KPIs: {kpi_list}")
    print(f"Cars: {car_list}")
except (FileNotFoundError, json.JSONDecodeError, KeyError) as e:
    print(f"⚠️ {input_filename} not found, invalid, or doesn't contain 20 cars. Fetching new data... ({str(e)})")
    # Fetch top 10 ICE and 10 EV cars with model tiers for 2024
    prompt_cars = "List the 10 most bought Internal Combustion Engine (ICE) cars and the 10 most bought Electric Vehicles (EV) in the US for 2024 as a simple list of names with specific model tiers (e.g., 'Toyota Camry LE', 'Tesla Model 3 Long Range'), one per line, no numbers or extra text."
    response_cars = query_x_api(prompt_cars)
    if response_cars and "choices" in response_cars and len(response_cars["choices"]) > 0:
        car_list = response_cars["choices"][0]["message"]["content"].strip().split('\n')
        car_list = [car.strip() for car in car_list if car.strip()]  # Clean up the list
        if len(car_list) != 20:
            print(f"❌ Fetched {len(car_list)} cars instead of 20. Exiting due to insufficient data.")
            sys.exit(1)
    else:
        print("❌ Failed to fetch car list from API. Exiting due to insufficient data.")
        sys.exit(1)
    
    # Define a pool of potential KPIs and exclude unwanted ones
    all_kpis = [
        "Fuel Efficiency (MPG)", "Acceleration (0-60 mph)", "Range (miles)", 
        "Maintenance Cost", "Cost Over Ownership", "Passenger Capacity", 
        "Cargo Space (cu ft)", "Towing Capacity (lbs)", "Horsepower"
    ]
    excluded_kpis = ["Reliability Rating", "Safety Rating", "Resale Value"]
    available_kpis = [kpi for kpi in all_kpis if kpi not in excluded_kpis]
    kpi_list = random.sample(available_kpis, 5)  # Randomly select 5 KPIs
    
    # Update data with the new car list and KPIs
    data = {"top_cars_US_2024": car_list, "top_5_KPIs": kpi_list}
    # Save the updated data to replace the missing or invalid file
    with open(input_filename, "w") as json_file:
        json.dump(data, json_file, indent=4)
    print(f"✅ Created/Updated {input_filename} with 2024 data (10 ICE + 10 EV) with tiers and new KPIs: {kpi_list}")

# Step 2: Fetch KPI data for each car for 2024
car_kpi_data = {}
for car in car_list:
    car_kpi_data[car] = {}
    for kpi in kpi_list:
        # Define "Cost Over Ownership" as 5-year total cost in thousands of dollars
        if kpi == "Cost Over Ownership":
            prompt = f"Provide the 5-year total cost of ownership in thousands of dollars for the {car} in the US for 2024 as a single integer value (e.g., 25 for $25,000, no units, no text, just the number)."
        else:
            prompt = f"Provide the {kpi} for the {car} in the US for 2024 as a single integer value (no units, no text, just the number)."
        response = query_x_api(prompt)
        value = extract_content(response)
        car_kpi_data[car][kpi] = value

# Step 3: Calculate averages for each KPI and add as a new entry
car_kpi_data["Average"] = {}
for kpi in kpi_list:
    kpi_values = [car_kpi_data[car][kpi] for car in car_list if car_kpi_data[car][kpi] is not None]
    avg_value = statistics.mean(kpi_values) if kpi_values else 0
    car_kpi_data["Average"][kpi] = avg_value
    print(f"📊 Calculated average for {kpi}: {avg_value} (from values {kpi_values})")

# Step 4: Save the new data to a JSON file
output_filename = "car_kpi_values_2024.json"
try:
    with open(output_filename, "w") as json_file:
        json.dump(car_kpi_data, json_file, indent=4)
    print(f"✅ Data saved to {output_filename}")
except Exception as e:
    print(f"❌ Error saving to {output_filename}: {str(e)}")