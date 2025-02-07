import os
from flask import Flask, request, jsonify
from datetime import datetime

app = Flask(__name__)

# Get API key from environment variable
API_KEY = os.getenv("API_KEY")

def log_weather_delay(order_data):
    """Log weather delays in a file."""
    log_message = f"{datetime.now()} - Delay Alert: {order_data['order_id']} - {order_data['location']} - {order_data['weather_condition']}\n"
    
    # Save log to a file
    with open("weather_delays.log", "a") as log_file:
        log_file.write(log_message)
    
    print(log_message)  # Also print in Render logs

@app.route('/')
def home():
    return "Weather Alert / Client Delayed Orders Service is Running!"

@app.route('/check_orders', methods=['POST'])
def check_orders():
    provided_api_key = request.headers.get("X-API-KEY")

    if provided_api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 403

    data = request.json

    # If there's a weather delay, log it
    if "delay" in data.get("weather_condition", "").lower():
        log_weather_delay(data)

    return jsonify({"message": "Checked orders", "data_received": data})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
