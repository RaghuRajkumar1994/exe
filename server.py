from flask import Flask, render_template_string, redirect, url_for, Response, request
from flask_socketio import SocketIO, join_room, leave_room
import eventlet
from datetime import datetime
import pandas as pd
from collections import defaultdict
import io
import csv

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
# Dictionary to track which client (sid) belongs to which machine
SID_TO_MACHINE = {}

# --- NEW: Function to broadcast the online machine list ---
def broadcast_online_status():
    """Calculates the list of unique online machines and broadcasts it to all clients."""
    # Get a unique set of all machine names currently connected
    online_machines = list(set(SID_TO_MACHINE.values()))
    
    print(f"Currently online machines: {online_machines}")
    # The dashboard listens for 'update_machine_status'
    socketio.emit('update_machine_status', {'onlineMachines': online_machines})

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
        
        # Ensure quantity is treated as an integer for summation
        try:
             qty = int(entry['produced_qty'])
        except ValueError:
             qty = 0
             
        machine_qty_totals[entry['machine_name']] += qty

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
        
        plan_data_raw = df.head(10).to_dict('records')
        plan_data_processed = []
        for i, item in enumerate(plan_data_raw):
            item['line_id'] = f"{target_machine}_{i+1}" 
            item['status'] = 'pending' # Default status
            plan_data_processed.append(item)
        
        MACHINE_PLANS[target_machine] = plan_data_processed
        
        # Broadcast the new, correctly formatted plan
        socketio.emit('update_worker_plan', {'plan': plan_data_processed, 'machineName': target_machine}, room=target_machine)

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
        'FG Part Number', 'Cable Identification', 'Produced Qty', 
        'Produced Length', 'QTY PRODUCED HOURS'
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

@socketio.on('submit_output')
def handle_submit_output(data):
    """Handles production data submission from the worker interface."""
    try:
        # This line expects 'entry_date' and 'entry_time' keys
        data['datetime'] = datetime.strptime(f"{data['entry_date']} {data['entry_time']}", "%Y-%m-%d %H:%M") 
        data['produced_qty'] = data['produced_qty']
        data['produced_length'] = data['produced_length']
        data['qty_produced_hours'] = data['qty_produced_hours']
        
        SUBMISSION_LOG.append(data)
        
        # Broadcast the updated log to the dashboard
        broadcast_data(data['datetime'].strftime('%Y-%m-%d')) 
        
        # Acknowledge success to the worker (sender only)
        socketio.emit('submission_success', {'success': True}, room=request.sid)

    except KeyError as e:
        print(f"Error processing submission: {e}")
        socketio.emit('submission_success', {'success': False, 'reason': f"Missing data field: {e}"}, room=request.sid)
    except ValueError as e:
        print(f"Date/Time format error or invalid number: {e}")
        socketio.emit('submission_success', {'success': False, 'reason': f"Invalid data format: {e}"}, room=request.sid)

@socketio.on('join_machine_room')
def handle_join_machine_room(data):
    """Handles a worker joining a machine-specific room."""
    machine_name = data.get('machineName')

    if not machine_name:
        socketio.emit('join_confirm', {'success': False, 'reason': 'Missing machine name.'}, room=request.sid)
        return

    # 1. Update machine tracking
    # IMPORTANT: Use SID_TO_MACHINE.pop(request.sid) to handle re-joins if you wanted to be strict.
    # For simplicity, we just update it.
    SID_TO_MACHINE[request.sid] = machine_name
    
    # 2. Join the SocketIO room
    join_room(machine_name)
    print(f"Client {request.sid} joined room: {machine_name}")
    
    # 3. Broadcast the online status update (to dashboard)
    broadcast_online_status()

    # 4. Send the current plan back to the worker
    current_plan = MACHINE_PLANS.get(machine_name, [])
    socketio.emit('update_worker_plan', {'plan': current_plan, 'machineName': machine_name}, room=request.sid)

    # 5. Confirm success back to the worker (sender only)
    socketio.emit('join_confirm', {'success': True, 'machineName': machine_name}, room=request.sid)


@socketio.on('mark_plan_complete')
def handle_mark_plan_complete(data):
    """Marks a specific line in the plan sheet as complete for a machine."""
    line_id = data.get('lineId')
    machine_name = data.get('machineName')

    if not line_id or not machine_name:
        return

    if machine_name in MACHINE_PLANS:
        plan = MACHINE_PLANS[machine_name]
        for item in plan:
            if item.get('line_id') == line_id:
                item['status'] = 'completed'
                break
        
        # Broadcast the updated plan back to all clients in that machine's room
        socketio.emit('update_worker_plan', {'plan': plan, 'machineName': machine_name}, room=machine_name)

@socketio.on('send_live_message')
def handle_send_live_message(data):
    """Sends a live message from the dashboard to a specific machine room."""
    target_machine = data.get('targetMachine')
    message_text = data.get('messageText')

    if not target_machine or not message_text:
        socketio.emit('message_sent_confirm', {'success': False, 'machineName': target_machine, 'reason': 'Missing target machine or message text.'}, room=request.sid)
        return
    
    # Check if any client is connected to the target room
    # Note: Flask-SocketIO doesn't easily expose this, but we can check our internal map
    is_online = target_machine in SID_TO_MACHINE.values()

    if is_online:
        # Emit the message to all clients in the target room
        socketio.emit('live_message', {'message': message_text}, room=target_machine)
        socketio.emit('message_sent_confirm', {'success': True, 'machineName': target_machine}, room=request.sid)
    else:
        # If the machine is not online according to our map
         socketio.emit('message_sent_confirm', {'success': False, 'machineName': target_machine, 'reason': 'Machine is currently offline or not connected.'}, room=request.sid)


@socketio.on('request_dashboard_data')
def handle_request_dashboard_data(data):
    """Sends data to the dashboard based on a date filter request."""
    date_to_filter = data.get('date')
    broadcast_data(date_to_filter)


@socketio.on('connect')
def handle_connect():
    """Sends the current data when a client connects."""
    if request.path == '/dashboard':
        # Dashboard connects: send log data
        broadcast_data(date_str=datetime.now().strftime('%Y-%m-%d')) 
        # Dashboard connects: send current machine status (NEW)
        broadcast_online_status()

# -------------------------------------
# NEW Socket Event: Disconnect Handler (Crucial for online/offline tracking)
# -------------------------------------
@socketio.on('disconnect')
def handle_disconnect():
    """Removes the disconnected client from the SID_TO_MACHINE map and updates status."""
    if request.sid in SID_TO_MACHINE:
        machine_name = SID_TO_MACHINE.pop(request.sid)
        print(f"Client {request.sid} disconnected from room: {machine_name}")
        
        # IMPORTANT: Update the status on the dashboard since a machine may have gone offline
        broadcast_online_status()

# --- Start the Server ---
if __name__ == '__main__':
    print("Starting Flask-SocketIO Server...")
    print(f"Worker Input Page: http://10.10.2.230:5000/worker (e.g., set this to your worker tablet's home screen)")
    print(f"Dashboard Page: http://10.10.2.230:5000/dashboard")
    # Use eventlet.wsgi.server for production/async mode
    eventlet.wsgi.server(eventlet.listen(('', 5000)), app)