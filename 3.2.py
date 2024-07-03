from flask import Flask, render_template
import mysql.connector

app = Flask(__name__)

# กำหนดค่าการเชื่อมต่อกับฐานข้อมูล
db_config = {
    'host': '192.168.101.44',
    'user': 'madhan',
    'password': 'Bluefalo2012',
    'database': 'plc_data'
}

def get_data():
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, screw_feeder_speed, downsprout_temp, temp_zone1, temp_zone2, temp_zone3, feed_rate, extruder_load, extruder_speed, knife_load, pre_cond_steam, pre_cond_water, timestamp FROM plc_data")
    data = cursor.fetchall()
    cursor.close()
    conn.close()
    return data

@app.route('/')
def index():
    data = get_data()
    return render_template('index.html', data=data)

if __name__ == '__main__':
    app.run(debug=True)
