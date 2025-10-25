from flask import Flask, render_template_string, redirect, url_for, Response, request
from flask_socketio import SocketIO
import eventlet
from datetime import datetime
import pandas as pd
from collections import defaultdict

# --- Configuration ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your_super_secure_secret_key' 
socketio = SocketIO(app, async_mode='eventlet')

# --- Data Storage (In-memory list of all submissions) ---
SUBMISSION_LOG = [] 

# --- Helper Function to Get Data by Date ---
def get_data_for_date(date_str):
    """Filters the submission log for a specific date (YYYY-MM-DD)."""
    if not date_str:
        return SUBMISSION_LOG

    try:
        filter_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    except ValueError:
        return SUBMISSION_LOG

    filtered_log = [
        entry for entry in SUBMISSION_LOG 
        if entry['datetime'].date() == filter_date
    ]
    
    filtered_log.sort(key=lambda x: x['datetime'], reverse=True) 
    return filtered_log

# --- Broadcast Function ---
def broadcast_data(date_str=None):
    """
    Broadcasts the data for the requested date to the dashboard.
    Includes aggregation for charts.
    """
    if not date_str:
        date_str = datetime.now().strftime('%Y-%m-%d')
        
    log_to_send = get_data_for_date(date_str)
    
    data_to_send = []
    machine_qty_totals = defaultdict(int) 

    for entry in log_to_send:
        # 1. Prepare raw data for the table
        clean_entry = {
            'time_display': entry['datetime'].strftime("%Y-%m-%d %H:%M:%S"),
            'worker_name': entry['worker_name'],
            'shift': entry['shift'], 
            'machine_name': entry['machine_name'],
            'order_no': entry['order_no'],
            'fg_part_no': entry['fg_part_no'],
            'applicator_no': entry['applicator_no'],
            'cable_id': entry['cable_id'],
            'produced_qty': entry['produced_qty'],
            'produced_length': entry['produced_length'],
            'worked_hours': entry['worked_hours']
        }
        data_to_send.append(clean_entry)
        
        # 2. Aggregate data for the chart (Qty by Machine)
        machine_qty_totals[entry['machine_name']] += entry['produced_qty']

    # Convert defaultdict to a list of objects for easier JS consumption
    chart_data = [{'machine': k, 'total_qty': v} for k, v in machine_qty_totals.items()]

    data = {
        'log': data_to_send,
        'chart_data': chart_data # New payload for the chart
    }
    socketio.emit('update_dashboard', data) 
    
# --- Flask Routes ---

@app.route('/worker')
def worker_page():
    """Serves the worker input interface."""
    try:
        with open('worker.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        return render_template_string(html_content)
    except FileNotFoundError:
        return "Error: worker.html not found. Ensure it is in the same directory.", 404

@app.route('/dashboard')
def dashboard_page():
    """Serves the live dashboard interface."""
    try:
        with open('dashboard.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        return render_template_string(html_content)
    except FileNotFoundError:
        return "Error: dashboard.html not found. Ensure it is in the same directory.", 404

@app.route('/')
def index():
    return redirect(url_for('dashboard_page'))

# -------------------------------------
# Export Route (/export)
# -------------------------------------
@app.route('/export', methods=['GET'])
def export_data():
    """Exports all stored data to a CSV file."""
    
    FIELD_NAMES = [
        'Date/Time', 'Shift', 'Worker Name', 'Machine Name', 'Order No', 
        'FG Part Number', 'Applicator No', 'Cable Identification', 
        'Produced Qty', 'Produced Length', 'Worked Hours'
    ]
    
    rows = []
    sorted_log = sorted(SUBMISSION_LOG, key=lambda x: x['datetime'])

    for entry in sorted_log:
        rows.append({
            'Date/Time': entry['datetime'].strftime("%Y-%m-%d %H:%M:%S"),
            'Shift': entry['shift'], 
            'Worker Name': entry['worker_name'],
            'Machine Name': entry['machine_name'],
            'Order No': entry['order_no'],
            'FG Part Number': entry['fg_part_no'],
            'Applicator No': entry['applicator_no'],
            'Cable Identification': entry['cable_id'],
            'Produced Qty': entry['produced_qty'],
            'Produced Length': entry['produced_length'],
            'Worked Hours': entry['worked_hours']
        })

    df = pd.DataFrame(rows, columns=FIELD_NAMES)

    csv_data = df.to_csv(index=False, encoding='utf-8-sig')
    
    response = Response(
        csv_data,
        mimetype="text/csv",
        headers={
            "Content-disposition": "attachment; filename=production_data.csv",
            "Cache-Control": "no-cache"
        }
    )
    return response


# --- Socket.IO Event Handler ---

@socketio.on('submit_output')
def handle_output_submission(data):
    """Handles new output submissions with detailed fields."""
    global SUBMISSION_LOG
    
    entry_date = data.get('entryDate')
    entry_time = data.get('entryTime')
    shift = data.get('shift', 'N/A')
    
    worker_name = data.get('workerName', 'N/A')
    machine_name = data.get('machineName', 'N/A')
    order_no = data.get('orderNo', 'N/A')
    fg_part_no = data.get('fgPartNo', 'N/A')
    applicator_no = data.get('applicatorNo', 'N/A')
    cable_id = data.get('cableId', 'N/A')
    produced_qty_str = data.get('producedQty')
    produced_length_str = data.get('producedLength')
    worked_hours_str = data.get('workedHours')
    
    # Validation and conversion
    try:
        entry_datetime = datetime.strptime(f"{entry_date} {entry_time}", "%Y-%m-%d %H:%M")
        
        produced_qty_val = int(produced_qty_str)
        produced_length_val = float(produced_length_str)
        worked_hours_val = float(worked_hours_str)
        
        # Combined validation for all three numeric fields
        if produced_qty_val <= 0 or produced_length_val <= 0 or worked_hours_val <= 0: return 
    except Exception as e:
        print(f"Error converting data: {e}")
        return 

    # Create the log entry
    new_entry = {
        'datetime': entry_datetime,
        'worker_name': worker_name,
        'shift': shift,
        'machine_name': machine_name,
        'order_no': order_no,
        'fg_part_no': fg_part_no,
        'applicator_no': applicator_no,
        'cable_id': cable_id,
        'produced_qty': produced_qty_val,
        'produced_length': produced_length_val,
        'worked_hours': worked_hours_val
    }
    
    SUBMISSION_LOG.insert(0, new_entry)

    # Broadcast the updated data, filtering for the date just submitted
    broadcast_data(date_str=entry_datetime.strftime('%Y-%m-%d'))


@socketio.on('request_date_data')
def handle_date_request(data):
    """Handles explicit requests from the dashboard date filter."""
    date_to_filter = data.get('date')
    broadcast_data(date_to_filter)


@socketio.on('connect')
def handle_connect():
    """Sends the current data when a client connects."""
    broadcast_data(date_str=datetime.now().strftime('%Y-%m-%d')) 

# --- Start the Server ---

if __name__ == '__main__':
    print("Starting Flask-SocketIO Server...")
    print(f"Worker Input Page: http://127.0.0.1:5000/worker")
    print(f"Live Dashboard: http://127.0.0.1:5000/dashboard")
    socketio.run(app, port=5000)