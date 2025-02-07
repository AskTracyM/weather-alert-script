from flask import Flask, request, jsonify

app = Flask(__name__)

# Define a SECRET API KEY
API_KEY = "EpLN&6zX"  # Change this to a strong, random key!

@app.route('/')
def home():
    return "Weather Alert / Client Delayed Orders Service is Running!"

@app.route('/check_orders', methods=['POST'])
def check_orders():
    # Get API key from the request header
    provided_api_key = request.headers.get("X-API-KEY")

    # Verify API key
    if provided_api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 403  # 403 = Forbidden

    # If key is valid, process the request
    data = request.json
    return jsonify({"message": "Checked orders", "data_received": data})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
