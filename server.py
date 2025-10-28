from flask import Flask, render_template_string, redirect, url_for, Response, request
from flask_socketio import SocketIO, join_room
import eventlet
from datetime import datetime
import pandas as pd
from collections import defaultdict
import io

try:
    import openpyxl
except ImportError:
    print("Warning: openpyxl is not installed. Excel file uploads will fail.")


# --- Configuration ---
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024 
app.config['SECRET_KEY'] = 'your_super_secure_secret_key' 
socketio = SocketIO(app, async_mode='eventlet')

# --- Data Storage (In-memory) ---
SUBMISSION_LOG = [] 
MACHINE_PLANS = {} 
SID_TO_MACHINE = {}

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
    """
    if not date_str:
        date_str = datetime.now().strftime('%Y-%m-%d')
        
    log_to_send = get_data_for_date(date_str)
    
    data_to_send = []
    machine_qty_totals = defaultdict(int) 

    for entry in log_to_send:
        clean_entry = {
            'time_display': entry['datetime'].strftime("%Y-%m-%d %H:%M:%S"),
            'worker_name': entry['worker_name'],
            'shift': entry['shift'], 
            'machine_name': entry['machine_name'],
            'fg_part_no': entry['fg_part_no'],
            'cable_id': entry['cable_id'],
            'produced_qty': entry['produced_qty'],
            'produced_length': entry['produced_length'],
            'qty_produced_hours': entry['qty_produced_hours']
        }
        data_to_send.append(clean_entry)
        
        machine_qty_totals[entry['machine_name']] += entry['produced_qty']

    chart_data = [{'machine': k, 'total_qty': v} for k, v in machine_qty_totals.items()]

    data = {
        'log': data_to_send,
        'chart_data': chart_data,
        'machines': sorted(list(MACHINE_PLANS.keys())) 
    }
    socketio.emit('update_dashboard', data) 
    
# --- Flask Routes ---

@app.route('/worker')
def worker_page():
    """Serves the worker input interface."""
    try:
        # Use existing worker.html content
        with open('worker.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        return render_template_string(html_content)
    except FileNotFoundError:
        return "Error: worker.html not found. Ensure it is in the same directory.", 404

@app.route('/dashboard')
def dashboard_page():
    """Serves the live dashboard interface."""
    try:
        # Use existing dashboard.html content
        with open('dashboard.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        return render_template_string(html_content)
    except FileNotFoundError:
        return "Error: dashboard.html not found. Ensure it is in the same directory.", 404

@app.route('/')
def index():
    return redirect(url_for('dashboard_page'))

@app.route('/upload_plan', methods=['POST'])
def upload_plan():
    """Handles Excel file upload, reads data, and broadcasts the plan to the target machine."""
    
    target_machine = request.form.get('target_machine')
    excel_file = request.files.get('plan_sheet')

    if not target_machine or not excel_file:
        return "Error: Missing machine name or file.", 400

    if not excel_file.filename.endswith(('.xlsx', '.xls')):
        return "Error: Invalid file format. Please upload an Excel file (.xlsx or .xls).", 400

    try:
        file_stream = io.BytesIO(excel_file.read())
        df = pd.read_excel(file_stream, sheet_name=0)
        df = df.astype(str)
        plan_data = df.head(10).to_dict('records')

        MACHINE_PLANS[target_machine] = plan_data
        
        socketio.emit('update_worker_plan', {'plan': plan_data, 'machineName': target_machine}, room=target_machine)

        broadcast_data(datetime.now().strftime('%Y-%m-%d'))
        
        return f"Success: Plan sheet for {target_machine} uploaded and sent to machine room.", 200

    except ImportError:
        return "Error processing file: Missing dependency 'openpyxl'. Please install it to enable Excel reading.", 500
    except Exception as e:
        print(f"File processing error: {e}")
        return f"Error processing file: {str(e)}", 500

@app.route('/export', methods=['GET'])
def export_data():
    """Exports all stored production data to a CSV file."""
    
    FIELD_NAMES = [
        'Date/Time', 'Shift', 'Worker Name', 'Machine Name', 
        'FG Part Number', 'Cable Identification', 
        'Produced Qty', 'Produced Length', 'QTY PRODUCED HOURS'
    ]
    
    rows = []
    sorted_log = sorted(SUBMISSION_LOG, key=lambda x: x['datetime'])

    for entry in sorted_log:
        rows.append({
            'Date/Time': entry['datetime'].strftime("%Y-%m-%d %H:%M:%S"),
            'Shift': entry['shift'], 
            'Worker Name': entry['worker_name'],
            'Machine Name': entry['machine_name'],
            'FG Part Number': entry['fg_part_no'],
            'Cable Identification': entry['cable_id'],
            'Produced Qty': entry['produced_qty'],
            'Produced Length': entry['produced_length'],
            'QTY PRODUCED HOURS': entry['qty_produced_hours']
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


# --- Socket.IO Event Handlers ---

@socketio.on('join_machine_room')
def handle_join_machine_room(data):
    """Worker joins a room based on their machine name."""
    machine_name = data.get('machineName')
    
    if machine_name:
        join_room(machine_name)
        SID_TO_MACHINE[request.sid] = machine_name
        print(f"Client {request.sid} joined room: {machine_name}")
        
        if machine_name in MACHINE_PLANS:
            plan_data = MACHINE_PLANS[machine_name]
            socketio.emit('update_worker_plan', {'plan': plan_data, 'machineName': machine_name}, room=request.sid)
        else:
            socketio.emit('update_worker_plan', {'plan': [], 'machineName': machine_name, 'message': 'No active plan found.'}, room=request.sid)

# -------------------------------------
# New Socket Event: Send Message to Machine
# -------------------------------------
@socketio.on('send_message_to_machine')
def handle_send_message(data):
    """Sends a text message to a specific machine's room."""
    machine_name = data.get('machineName')
    message_text = data.get('message')
    sender = data.get('sender', 'System')
    
    if not machine_name or not message_text:
        # Send failure confirmation back to the sender (dashboard)
        socketio.emit('message_sent_confirm', {
            'success': False, 
            'machineName': machine_name, 
            'reason': 'Missing machine name or message text'
        }, room=request.sid)
        return

    # 1. Broadcast the message to the specific machine room (room=machine_name)
    message_data = {
        'sender': sender,
        'message': message_text
    }
    
    socketio.emit('receive_message', message_data, room=machine_name)
    
    # 2. Send success confirmation back to the sender (dashboard)
    socketio.emit('message_sent_confirm', {
        'success': True, 
        'machineName': machine_name
    }, room=request.sid)

@socketio.on('submit_output')
def handle_output_submission(data):
    """Handles new output submissions with detailed fields."""
    global SUBMISSION_LOG
    
    entry_datetime = datetime.strptime(f"{data.get('entryDate')} {data.get('entryTime')}", "%Y-%m-%d %H:%M")
    shift = data.get('shift', 'N/A')
    worker_name = data.get('workerName', 'N/A')
    machine_name = data.get('machineName', 'N/A')
    fg_part_no = data.get('fgPartNo', 'N/A')
    cable_id = data.get('cableId', 'N/A')
    
    try:
        produced_qty_val = int(data.get('producedQty'))
        produced_length_val = float(data.get('producedLength'))
        qty_produced_hours_val = float(data.get('qtyProducedHours')) 
        
        if produced_qty_val <= 0 or produced_length_val <= 0 or qty_produced_hours_val <= 0: return 
    except Exception as e:
        print(f"Error converting data: {e}")
        return 

    new_entry = {
        'datetime': entry_datetime,
        'worker_name': worker_name,
        'shift': shift,
        'machine_name': machine_name,
        'fg_part_no': fg_part_no,
        'cable_id': cable_id,
        'produced_qty': produced_qty_val,
        'produced_length': produced_length_val,
        'qty_produced_hours': qty_produced_hours_val
    }
    
    SUBMISSION_LOG.insert(0, new_entry)

    broadcast_data(date_str=entry_datetime.strftime('%Y-%m-%d'))


@socketio.on('request_date_data')
def handle_date_request(data):
    """Handles explicit requests from the dashboard date filter."""
    date_to_filter = data.get('date')
    broadcast_data(date_to_filter)


@socketio.on('connect')
def handle_connect():
    """Sends the current data when a client connects."""
    if request.path == '/dashboard':
        broadcast_data(date_str=datetime.now().strftime('%Y-%m-%d')) 

# --- Start the Server ---

if __name__ == '__main__':
    print("Starting Flask-SocketIO Server...")
    # NOTE: host='0.0.0.0' allows access from other devices on your local network (LAN)
    print(f"Worker Input Page: http://10.10.2.230:5000/worker (e.g., http://<YOUR_SERVER_IP>:5000/worker)")
    print(f"Live Dashboard: http://10.10.2.230:5000/dashboard (e.g., http://<YOUR_SERVER_IP>:5000/dashboard)")
    
    # If you see WinError 10048, change port=5000 to port=5001
    socketio.run(app, host='0.0.0.0', port=5000)