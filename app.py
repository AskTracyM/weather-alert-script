from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/')
def home():
    return "Weather Alert / Client Delayed Orders Service is Running!"

@app.route('/check_orders', methods=['POST'])
def check_orders():
    data = request.json
    # Your script logic goes here...
    response = {"message": "Checked orders", "data_received": data}
    return jsonify(response)

if __name__ == '__main__':
    app.run(debug=True)
