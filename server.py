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

    # Filter log entries that have the 'datetime' key and match the date
    filtered_log = [
        entry for entry in SUBMISSION_LOG 
        if entry.get('datetime') and entry['datetime'].date() == filter_date
    ]
    
    filtered_log.sort(key=lambda x: x['datetime'], reverse=True) 
    return filtered_log

# --- Broadcast Function (UPDATED for ALL Terminal Fields) ---
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
        # Note: The server combines Measured and Manual into a single field for display/export
        t1_crimp_height = entry.get('t1_crimp_height_manual') or entry.get('t1_crimp_height_measured')
        t1_insulation_height = entry.get('t1_insulation_height_manual') or entry.get('t1_insulation_height_measured')
        t1_crimp_width = entry.get('t1_crimp_width_manual') or entry.get('t1_crimp_width_measured')
        t1_insulation_width = entry.get('t1_insulation_width_manual') or entry.get('t1_insulation_width_measured')
        t1_pull_force = entry.get('t1_pull_force_manual') or entry.get('t1_pull_force_measured')
        
        t2_crimp_height = entry.get('t2_crimp_height_manual') or entry.get('t2_crimp_height_measured')
        t2_insulation_height = entry.get('t2_insulation_height_manual') or entry.get('t2_insulation_height_measured')
        t2_crimp_width = entry.get('t2_crimp_width_manual') or entry.get('t2_crimp_width_measured')
        t2_insulation_width = entry.get('t2_insulation_width_manual') or entry.get('t2_insulation_width_measured')
        t2_pull_force = entry.get('t2_pull_force_manual') or entry.get('t2_pull_force_measured')
        
        clean_entry = {
            'time_display': entry['datetime'].strftime("%Y-%m-%d %H:%M:%S"),
            'worker_name': entry.get('operator_name', 'N/A'), 
            'shift': entry.get('shift', 'N/A'), 
            'machine_name': entry.get('machine_name', 'N/A'),
            'fg_part_no': entry.get('fg_part_no', 'N/A'),
            'cable_id': entry.get('cable_id', 'N/A'),
            'produced_qty': entry.get('produced_qty', 0),
            'produced_length': entry.get('produced_length', 0.0),
            'qty_produced_hours': entry.get('qty_produced_hours', 0.0),
            # --- TERMINAL FIELDS ADDED FOR DASHBOARD ---
            't1_terminal_id': entry.get('t1_terminal_id', ''), # <-- T1 Part No
            't1_apl_no': entry.get('t1_apl_no', ''), 
            't1_crimp_height': t1_crimp_height,
            't1_insulation_height': t1_insulation_height,
            't1_crimp_width': t1_crimp_width,
            't1_insulation_width': t1_insulation_width,
            't1_pull_force': t1_pull_force,
            't2_terminal_id': entry.get('t2_terminal_id', ''), # <-- T2 Part No
            't2_apl_no': entry.get('t2_apl_no', ''), 
            't2_crimp_height': t2_crimp_height,
            't2_insulation_height': t2_insulation_height,
            't2_crimp_width': t2_crimp_width,
            't2_insulation_width': t2_insulation_width,
            't2_pull_force': t2_pull_force
            # ---------------------------------------------
        }
        data_to_send.append(clean_entry)
        
        # Ensure quantity is treated as an integer for summation
        try:
             qty = int(entry.get('produced_qty', 0))
        except (ValueError, TypeError):
             qty = 0
             
        machine_qty_totals[entry.get('machine_name', 'UNKNOWN')] += qty

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
        # Use header=0 to correctly read column names
        df = pd.read_excel(file_stream, sheet_name=0, header=0) 
        # Convert all relevant columns to string to prevent issues with mixed types
        df = df.fillna('').astype(str)
        
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

# -------------------------------------
# EXPORT DATA FUNCTION (UPDATED for ALL Terminal Fields)
# -------------------------------------
@app.route('/export', methods=['GET'])
def export_data():
    """Exports all stored production data to a CSV file, including terminal data."""
    
    rows = []
    sorted_log = sorted(SUBMISSION_LOG, key=lambda x: x.get('datetime', datetime.min)) 

    for entry in sorted_log:
        if 'datetime' not in entry:
            continue
            
        # Merge Measured and Manual data, preferring Manual if available
        t1_crimp_height = entry.get('t1_crimp_height_manual') or entry.get('t1_crimp_height_measured')
        t1_insulation_height = entry.get('t1_insulation_height_manual') or entry.get('t1_insulation_height_measured')
        t1_crimp_width = entry.get('t1_crimp_width_manual') or entry.get('t1_crimp_width_measured')
        t1_insulation_width = entry.get('t1_insulation_width_manual') or entry.get('t1_insulation_width_measured')
        t1_pull_force = entry.get('t1_pull_force_manual') or entry.get('t1_pull_force_measured')
        
        t2_crimp_height = entry.get('t2_crimp_height_manual') or entry.get('t2_crimp_height_measured')
        t2_insulation_height = entry.get('t2_insulation_height_manual') or entry.get('t2_insulation_height_measured')
        t2_crimp_width = entry.get('t2_crimp_width_manual') or entry.get('t2_crimp_width_measured')
        t2_insulation_width = entry.get('t2_insulation_width_manual') or entry.get('t2_insulation_width_measured')
        t2_pull_force = entry.get('t2_pull_force_manual') or entry.get('t2_pull_force_measured')


        row = {
            'datetime_obj': entry['datetime'],
            'Shift': entry.get('shift', ''),
            'Worker Name': entry.get('operator_name', ''), 
            'Machine Name': entry.get('machine_name', ''),
            'FG Part Number': entry.get('fg_part_no', ''),
            'Cable Identification': entry.get('cable_id', ''),
            'Produced Qty': entry.get('produced_qty', 0), 
            'Produced Length': entry.get('produced_length', 0.0),
            'QTY PRODUCED HOURS': entry.get('qty_produced_hours', 0.0),
            # --- TERMINAL FIELDS ADDED FOR EXPORT (Using the merged value) ---
            'T1 Part No': entry.get('t1_terminal_id', ''), # <-- T1 Part No
            'T1 APL NO': entry.get('t1_apl_no', ''), 
            'T1 Crimp H': t1_crimp_height,
            'T1 Insul H': t1_insulation_height,
            'T1 Crimp W': t1_crimp_width,
            'T1 Insul W': t1_insulation_width,
            'T1 Pull F (N)': t1_pull_force,
            'T2 Part No': entry.get('t2_terminal_id', ''), # <-- T2 Part No
            'T2 APL NO': entry.get('t2_apl_no', ''), 
            'T2 Crimp H': t2_crimp_height,
            'T2 Insul H': t2_insulation_height,
            'T2 Crimp W': t2_crimp_width,
            'T2 Insul W': t2_insulation_width,
            'T2 Pull F (N)': t2_pull_force
            # ---------------------------------------------
        }
        rows.append(row)

    if not rows:
        return "No data to export", 204
        
    df = pd.DataFrame(rows)

    # 2. Split the datetime_obj into separate 'Date' and 'Time' columns
    if 'datetime_obj' in df.columns:
        # Create new 'Date' and 'Time' columns at the start of the DataFrame
        df.insert(0, 'Date', df['datetime_obj'].dt.strftime('%Y-%m-%d'))
        df.insert(1, 'Time', df['datetime_obj'].dt.strftime('%H:%M:%S'))
        
        # Remove the original datetime object column
        df = df.drop(columns=['datetime_obj'])

    # 3. Define and reorder columns (including new terminal fields)
    NEW_FIELD_NAMES = [
        'Date', 'Time', 'Shift', 'Worker Name', 'Machine Name', 
        'FG Part Number', 'Cable Identification', 'Produced Qty', 
        'Produced Length', 'QTY PRODUCED HOURS',
        'T1 Part No', 'T1 APL NO', 'T1 Crimp H', 'T1 Insul H', 'T1 Crimp W', 'T1 Insul W', 'T1 Pull F (N)', 
        'T2 Part No', 'T2 APL NO', 'T2 Crimp H', 'T2 Insul H', 'T2 Crimp W', 'T2 Insul W', 'T2 Pull F (N)' 
    ]
    
    # Filter the final columns to only include those that actually exist 
    final_cols = [col for col in NEW_FIELD_NAMES if col in df.columns]
    df = df[final_cols]
    
    # 4. Generate CSV
    csv_data = df.to_csv(index=False, encoding='utf-8-sig', quoting=csv.QUOTE_ALL)

    response = Response(
        csv_data,
        mimetype="text/csv",
        headers={
            "Content-disposition": "attachment; filename=production_data_export.csv",
            "Cache-Control": "no-cache"
        }
    )
    return response

# --- Socket.IO Event Handlers ---

@socketio.on('submit_output')
def handle_submit_output(data):
    """Handles production data submission from the worker interface."""
    try:
        # Create a single datetime object from the date and time strings
        data['datetime'] = datetime.strptime(f"{data['entry_date']} {data['entry_time']}", "%Y-%m-%d %H:%M") 
        
        # Log the submitted data, including all terminal and measured fields
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
    
    # Check if any client is connected to the target room by checking our internal map
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
# Socket Event: Disconnect Handler (Crucial for online/offline tracking)
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
    print(f"Worker Input Page: http://0.0.0.0:5000/worker")
    print(f"Dashboard Page: http://0.0.0.0:5000/dashboard")
    # Use eventlet.wsgi.server for production/async mode
    eventlet.wsgi.server(eventlet.listen(('', 5000)), app)