import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import uuid
import datetime
import time
import threading
import webbrowser
from address_parser import parse_address_string, split_addresses
from ups_api import UPSApiClient
import openpyxl
from openpyxl.styles import Font

class UPSPickupGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("UPS Pickup Automation")
        self.root.geometry("800x800")
        
        self.api_client = UPSApiClient()
        self.stop_batch = False
        self.history_file = "pickup_history.json"
        self.manual_mode = tk.BooleanVar(value=True)
        self.next_step_event = threading.Event()
        
        # Trace to auto-resume when unchecking Manual Mode
        self.manual_mode.trace_add("write", lambda *args: self.next_step_event.set() if not self.manual_mode.get() else None)
        
        self.setup_ui()

    def log_message(self, message):
        """Simplistic logger for the GUI."""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}")

    def setup_ui(self):
        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 10))
        style.configure("TButton", font=("Helvetica", 10, "bold"))
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.single_tab = ttk.Frame(self.notebook)
        self.batch_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.single_tab, text="Single Pickup")
        self.notebook.add(self.batch_tab, text="Batch Pickup")
        
        self.setup_single_tab()
        self.setup_batch_tab()

    def setup_single_tab(self):
        main_frame = ttk.Frame(self.single_tab, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Address Paste Area
        ttk.Label(main_frame, text="Paste Address Here:").pack(anchor=tk.W)
        self.address_text = tk.Text(main_frame, height=5, font=("Helvetica", 10))
        self.address_text.pack(fill=tk.X, pady=(0, 10))
        
        parse_btn = ttk.Button(main_frame, text="Parse Address", command=self.parse_and_fill)
        parse_btn.pack(pady=(0, 20))
        
        # Auditable Fields
        fields_frame = ttk.LabelFrame(main_frame, text="Audit Details", padding="10")
        fields_frame.pack(fill=tk.BOTH, expand=True)
        
        self.fields = {}
        field_labels = ["Street", "Suite", "City", "State", "Zip", "Country", "CompanyName", "ContactName", "Phone", "Email"]
        
        for i, label in enumerate(field_labels):
            ttk.Label(fields_frame, text=f"{label}:").grid(row=i, column=0, sticky=tk.W, pady=2)
            entry = ttk.Entry(fields_frame, width=40)
            entry.grid(row=i, column=1, sticky=tk.W, padx=10, pady=2)
            self.fields[label] = entry
            
        # Default values
        self.fields["Country"].insert(0, "CA")
        self.fields["CompanyName"].insert(0, "Omnitrans Inc")
        self.fields["ContactName"].insert(0, "Warehouse")
        
        # Date / Time Fields
        datetime_frame = ttk.LabelFrame(main_frame, text="Pickup Date & Time", padding="10")
        datetime_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(datetime_frame, text="Pickup Date:").grid(row=0, column=0, sticky=tk.W, pady=2)
        today = datetime.datetime.now()
        dates = [(today + datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(14)]
        self.date_cb = ttk.Combobox(datetime_frame, values=dates, state="readonly", width=12)
        self.date_cb.set(dates[0])
        self.date_cb.grid(row=0, column=1, sticky=tk.W, padx=10, pady=2)
        
        times = [f"{str(h).zfill(2)}:00" for h in range(8, 19)]
        ttk.Label(datetime_frame, text="Ready Time:").grid(row=0, column=2, sticky=tk.W, pady=2)
        self.ready_cb = ttk.Combobox(datetime_frame, values=times, state="readonly", width=8)
        self.ready_cb.set("09:00")
        self.ready_cb.grid(row=0, column=3, sticky=tk.W, padx=10, pady=2)
        
        ttk.Label(datetime_frame, text="Close Time:").grid(row=0, column=4, sticky=tk.W, pady=2)
        self.close_cb = ttk.Combobox(datetime_frame, values=times, state="readonly", width=8)
        self.close_cb.set("17:00")
        self.close_cb.grid(row=0, column=5, sticky=tk.W, padx=10, pady=2)
        
        # Action Buttons
        btn_frame = ttk.Frame(main_frame, padding="20")
        btn_frame.pack(fill=tk.X)
        
        submit_btn = ttk.Button(btn_frame, text="Generate Pickup", command=self.submit_pickup)
        submit_btn.pack(side=tk.RIGHT)
        
        cancel_btn = ttk.Button(btn_frame, text="Cancel Pickup", command=self.cancel_pickup_gui)
        cancel_btn.pack(side=tk.RIGHT, padx=10)
        
        clear_btn = ttk.Button(btn_frame, text="Clear", command=self.clear_fields)
        clear_btn.pack(side=tk.LEFT)
        
        history_btn = ttk.Button(btn_frame, text="View History", command=self.show_history_window)
        history_btn.pack(side=tk.LEFT, padx=10)

    def setup_batch_tab(self):
        main_frame = ttk.Frame(self.batch_tab, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Paste Multiple Addresses Here (Preferred: double newlines, but single lines also work):").pack(anchor=tk.W)
        self.batch_text = tk.Text(main_frame, height=10, font=("Helvetica", 10))
        self.batch_text.pack(fill=tk.X, pady=(0, 10))

        # Batch Defaults
        defaults_frame = ttk.LabelFrame(main_frame, text="Batch Default Values (Used if field missing in paste)", padding="10")
        defaults_frame.pack(fill=tk.X, pady=10)

        self.batch_defaults = {}
        batch_field_labels = ["CompanyName", "ContactName", "Phone", "Email", "ServiceCode"]
        for i, label in enumerate(batch_field_labels):
            ttk.Label(defaults_frame, text=f"{label}:").grid(row=i//2, column=(i%2)*2, sticky=tk.W, pady=2)
            entry = ttk.Entry(defaults_frame, width=20)
            entry.grid(row=i//2, column=(i%2)*2+1, sticky=tk.W, padx=10, pady=2)
            self.batch_defaults[label] = entry

        # Defaults set
        self.batch_defaults["CompanyName"].insert(0, "Omnitrans Inc")
        self.batch_defaults["ContactName"].insert(0, "Warehouse")
        self.batch_defaults["Phone"].insert(0, "5142886664")
        self.batch_defaults["ServiceCode"].insert(0, "011")

        # Batch Date/Time / Country
        time_country_frame = ttk.Frame(main_frame)
        time_country_frame.pack(fill=tk.X, pady=5)

        ttk.Label(time_country_frame, text="Pickup Date:").pack(side=tk.LEFT)
        today = datetime.datetime.now()
        dates = [(today + datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(14)]
        self.batch_date = ttk.Combobox(time_country_frame, values=dates, state="readonly", width=12)
        self.batch_date.set(dates[0])
        self.batch_date.pack(side=tk.LEFT, padx=5)
        ttk.Label(time_country_frame, text="Default Country:").pack(side=tk.LEFT)
        self.batch_country = ttk.Combobox(time_country_frame, values=["US", "CA"], width=5)
        self.batch_country.set("CA")
        self.batch_country.pack(side=tk.LEFT, padx=5)

        # Time selection
        time_frame = ttk.Frame(main_frame)
        time_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(time_frame, text="Ready Time:").pack(side=tk.LEFT)
        self.batch_ready_time = ttk.Combobox(time_frame, values=[f"{h:02d}00" for h in range(7, 21)], width=8)
        self.batch_ready_time.set("0900")
        self.batch_ready_time.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(time_frame, text="Close Time:").pack(side=tk.LEFT)
        self.batch_close_time = ttk.Combobox(time_frame, values=[f"{h:02d}00" for h in range(7, 23)], width=8)
        self.batch_close_time.set("1700")
        self.batch_close_time.pack(side=tk.LEFT, padx=5)

        # Progress Table
        ttk.Label(main_frame, text="Processing Progress:").pack(anchor=tk.W, pady=(10, 0))
        columns = ("Address", "Status", "PRN/Error")
        self.batch_tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=8)
        self.batch_tree.heading("Address", text="Address")
        self.batch_tree.column("Address", width=300)
        self.batch_tree.heading("Status", text="Status")
        self.batch_tree.column("Status", width=100)
        self.batch_tree.heading("PRN/Error", text="PRN / Error Detail")
        self.batch_tree.column("PRN/Error", width=300)
        self.batch_tree.pack(fill=tk.BOTH, expand=True, pady=5)

        # Mode controls
        mode_frame = ttk.Frame(main_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(mode_frame, text="Manual Mode (Step-by-Step)", variable=self.manual_mode).pack(side=tk.LEFT)
        self.next_btn = ttk.Button(mode_frame, text="Next Item >>", command=lambda: self.next_step_event.set(), state=tk.DISABLED)
        self.next_btn.pack(side=tk.LEFT, padx=10)

        # Buttons Frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)

        self.run_batch_btn = ttk.Button(btn_frame, text="Process Batch (Up to 150)", command=self.start_batch_thread)
        self.run_batch_btn.pack(side=tk.LEFT, padx=5)

        self.stop_batch_btn = ttk.Button(btn_frame, text="Stop Batch", command=self.stop_current_batch, state=tk.DISABLED)
        self.stop_batch_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="View History / Cancel", command=self.show_history_window).pack(side=tk.LEFT, padx=5)

    def parse_and_fill(self):
        raw_address = self.address_text.get("1.0", tk.END).strip()
        if not raw_address:
            messagebox.showwarning("Warning", "Please paste an address first.")
            return
            
        parsed = parse_address_string(raw_address)
        if parsed:
            for key in ["Street", "Suite", "City", "State", "Zip", "Phone", "Country", "CompanyName", "ContactName"]:
                if key in parsed and parsed[key]:
                    # Map CompanyName and ContactName to the fields dict keys if they differ
                    # In setup_ui, they are CompanyName and ContactName
                    if key in self.fields:
                        self.fields[key].delete(0, tk.END)
                        self.fields[key].insert(0, parsed[key])
        else:
            messagebox.showerror("Error", "Could not parse address. Please enter manually.")

    def select_all_history(self, tree):
        tree.selection_set(tree.get_children())

    def get_province_offset(self, state):
        """Returns the hour offset from EST for Canadian provinces."""
        offsets = {
            "NL": 1.5,
            "NS": 1, "NB": 1, "PE": 1,
            "QC": 0, "ON": 0,
            "MB": -1, "SK": -1,
            "AB": -2, "NT": -2,
            "BC": -3, "YT": -3
        }
        return offsets.get(state.upper(), 0)

    def adjust_time_for_timezone(self, pickup_data):
        """
        Ensures the ReadyTime is not in the past for same-day pickups.
        UPS expects the 'Local Time' of the pickup address.
        """
        if pickup_data.get("PickupDate") != datetime.datetime.now().strftime("%Y%m%d"):
            return # Only matters for same-day
            
        state = pickup_data.get("State", "ON")
        offset = self.get_province_offset(state)
        
        # Current local time at pickup location
        now_est = datetime.datetime.now()
        now_local = now_est + datetime.timedelta(hours=offset)
        local_time_str = now_local.strftime("%H%M")
        
        requested_ready = pickup_data.get("ReadyTime", "0900")
        
        # If requested ready time is in the past (or within 30 mins), push it forward
        if int(requested_ready) < int(local_time_str) + 30:
            # Shift to 1 hour from now, rounded up to the next hour
            new_ready = (now_local + datetime.timedelta(hours=1)).strftime("%H00")
            print(f"[Timezone] Safety Push for {state} (Currently {local_time_str} Local) from {requested_ready} to {new_ready}")
            pickup_data["ReadyTime"] = new_ready
            
            # Ensure CloseTime is still at least 2 hours after ReadyTime
            close_time = pickup_data.get("CloseTime", "1700")
            if int(close_time) < int(new_ready) + 200:
                pickup_data["CloseTime"] = str(int(new_ready) + 200).zfill(4)
                if int(pickup_data["CloseTime"]) > 2200: pickup_data["CloseTime"] = "2200"

    def submit_pickup(self):
        pickup_data = {k: v.get() for k, v in self.fields.items()}
        # Reverted back to the requested prepaid tracking number
        pickup_data["TrackingNumber"] = "1Z4A059A9190837115"
        
        pickup_data["PickupDate"] = self.date_cb.get().replace("-", "")
        pickup_data["ReadyTime"] = self.ready_cb.get().replace(":", "")
        pickup_data["CloseTime"] = self.close_cb.get().replace(":", "")
        
        if messagebox.askyesno("Audit Confirmation", f"Schedule pickup for {pickup_data['Street']}?"):
            try:
                # 1. Automate return label if no unique tracking is provided
                if pickup_data["TrackingNumber"] == "1Z4A059A9190837115":
                    # If it's the default, let's try to generate a unique one
                    label_res = self.api_client.create_return_label(pickup_data)
                    if label_res.get("status") == "success":
                        pickup_data["TrackingNumber"] = label_res["TrackingNumber"]
                    else:
                        print(f"[Warning] Automated label failed, using default: {label_res.get('message')}")

                # 2. Adjust for timezone if same-day
                self.adjust_time_for_timezone(pickup_data)

                # 3. Call the pickup API
                result = self.api_client.create_pickup(pickup_data)
                self.handle_api_result(result, pickup_data)
            except Exception as e:
                messagebox.showerror("API Error", str(e))

    def handle_api_result(self, result, pickup_data):
        prn = ""
        if "PickupCreationResponse" in result and "PRN" in result["PickupCreationResponse"]:
            prn = result["PickupCreationResponse"]["PRN"]
            self.show_success_dialog(prn, pickup_data)
        elif "response" in result and "errors" in result.get("response", {}):
             error_msg = json.dumps(result, indent=2)
             messagebox.showerror("API Error", f"Failed to schedule pickup:\n{error_msg}")
        else:
            msg = json.dumps(result, indent=2)
            messagebox.showinfo("API Response", f"Response details:\n{msg}")
            
        # Log to history
        full_addr = f"{pickup_data.get('Street', '')}, {pickup_data.get('City', '')}, {pickup_data.get('State', '')} {pickup_data.get('Zip', '')}, {pickup_data.get('Country', '')}".strip(", ")
        
        entry = {
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "company": pickup_data.get("CompanyName", ""),
            "address": full_addr,
            "date": pickup_data.get("PickupDate", ""),
            "status": "Success" if prn else "Failed",
            "prn": prn,
            "details": pickup_data # Save full data for robust reporting
        }
        self.save_to_history(entry)

    def show_success_dialog(self, prn, pickup_data):
        top = tk.Toplevel(self.root)
        top.title("Pickup Scheduled Successfully")
        top.geometry("400x350")
        top.transient(self.root)
        top.grab_set()

        msg = "Pickup Successfully Scheduled!\n\n"
        msg += f"PRN: {prn}\n"
        msg += f"Date: {pickup_data.get('PickupDate')}\n"
        msg += f"Time: {pickup_data.get('ReadyTime')} - {pickup_data.get('CloseTime')}\n"
        msg += f"Address: {pickup_data.get('Street')}\n\n"
        msg += "Please provide this information to the client."
        
        lbl = ttk.Label(top, text="Copy the information below:")
        lbl.pack(pady=(10, 0), padx=10, anchor=tk.W)
        
        txt = tk.Text(top, height=12, width=45, font=("Helvetica", 10))
        txt.insert(tk.END, msg)
        txt.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        btn = ttk.Button(top, text="Close", command=top.destroy)
        btn.pack(pady=(0, 10))

    def stop_current_batch(self):
        self.stop_batch = True
        
    def start_batch_thread(self):
        self.stop_batch = False
        raw_text = self.batch_text.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning("Warning", "Please paste addresses first.")
            return
            
        threading.Thread(target=self.process_batch, args=(raw_text,), daemon=True).start()

    def process_batch(self, raw_text):
        address_blocks = split_addresses(raw_text)
        if len(address_blocks) > 150:
            messagebox.showwarning("Limit Exceeded", f"Only first 150 addresses will be processed ({len(address_blocks)} found).")
            address_blocks = address_blocks[:150]

        # Clear previous progress
        for item in self.batch_tree.get_children():
            self.batch_tree.delete(item)

        defaults = {k: v.get() for k, v in self.batch_defaults.items()}
        common_date = self.batch_date.get().replace("-", "")
        default_country = self.batch_country.get()

        self.run_batch_btn.config(state=tk.DISABLED)
        self.stop_batch_btn.config(state=tk.NORMAL)
        
        ready_t = self.batch_ready_time.get()
        close_t = self.batch_close_time.get()
        
        for block in address_blocks:
            if self.stop_batch:
                self.log_message("Batch processing stopped by user.")
                break
                
            parsed = parse_address_string(block)
            if not parsed:
                self.batch_tree.insert("", tk.END, values=(block[:50]+"...", "Error", "Parsing Failed"))
                continue
                
            # Use batch default country if parser didn't find one or if it's US and batch default is CA
            if not parsed.get("Country") or (parsed.get("Country") == "US" and default_country == "CA"):
                 parsed["Country"] = default_country

            # Smart Service Code Selection based on Country
            service_code = defaults.get("ServiceCode")
            if parsed.get("Country") == "CA":
                service_code = "011" # UPS Standard for Canada
            elif parsed.get("Country") == "US":
                service_code = "003" # UPS Ground for US
            
            # Pre-validation
            missing = []
            if not parsed.get("Street"): missing.append("Street")
            if not parsed.get("City"): missing.append("City")
            if not parsed.get("State"): missing.append("State")
            if not parsed.get("Zip"): missing.append("Zip")
            
            if missing:
                self.batch_tree.insert("", tk.END, values=(block[:50]+"...", "Failed", f"Missing: {', '.join(missing)}"))
                continue

            # Merge with defaults
            pickup_data = {
                "Street": parsed.get("Street"),
                "Suite": parsed.get("Suite"),
                "City": parsed.get("City"),
                "State": parsed.get("State"),
                "Zip": parsed.get("Zip"),
                "Country": parsed.get("Country", "US"),
                "Phone": parsed.get("Phone") or defaults.get("Phone"),
                "CompanyName": defaults.get("CompanyName"),
                "ContactName": defaults.get("ContactName"),
                "ServiceCode": service_code,
                "Email": defaults.get("Email"),
                "PickupDate": common_date,
                "ReadyTime": ready_t,
                "CloseTime": close_t,
                "TrackingNumber": parsed.get("TrackingNumber") or "1Z4A059A9190837115"
            }

            iid = self.batch_tree.insert("", tk.END, values=(f"{pickup_data['Street']}, {pickup_data['City']}", "Processing...", ""))
            self.root.update_idletasks()

            try:
                # 1. Check if we need to generate a unique return label
                if not parsed.get("TrackingNumber"):
                    self.batch_tree.item(iid, values=(f"{pickup_data['Street']}, {pickup_data['City']}", "Creating Label...", ""))
                    label_res = self.api_client.create_return_label(pickup_data)
                    if label_res.get("status") == "success":
                        pickup_data["TrackingNumber"] = label_res["TrackingNumber"]
                    else:
                        err = label_res.get("message", "Label Error")
                        self.batch_tree.item(iid, values=(f"{pickup_data['Street']}, {pickup_data['City']}", "Failed (Label)", err))
                        continue

                # 2. Schedule the pickup
                self.adjust_time_for_timezone(pickup_data)
                self.batch_tree.item(iid, values=(f"{pickup_data['Street']}, {pickup_data['City']}", "Scheduling...", pickup_data["TrackingNumber"]))
                
                # Add a small delay for rate limiting
                time.sleep(1.5)
                
                result = self.api_client.create_pickup(pickup_data)
                prn = ""
                if "PickupCreationResponse" in result and "PRN" in result["PickupCreationResponse"]:
                    prn = result["PickupCreationResponse"]["PRN"]
                    self.log_message(f"Success for {pickup_data['Street']}: {prn}")
                    self.batch_tree.item(iid, values=(f"{pickup_data['Street']}, {pickup_data['City']}", "Success", prn))
                else:
                    err = result.get("response", {}).get("errors", [{}])[0].get("message", "API Error")
                    self.batch_tree.item(iid, values=(f"{pickup_data['Street']}, {pickup_data['City']}", "Failed", err))

                # Step-by-Step Pause Logic
                if self.manual_mode.get() and not self.stop_batch:
                    self.log_message("Manual Mode: Paused. Click 'Next Item' to continue.")
                    self.next_btn.config(state=tk.NORMAL)
                    self.next_step_event.wait()
                    self.next_step_event.clear()
                    self.next_btn.config(state=tk.DISABLED)
                
                # Log to history
                full_addr = f"{pickup_data.get('Street', '')}, {pickup_data.get('City', '')}, {pickup_data.get('State', '')} {pickup_data.get('Zip', '')}, {pickup_data.get('Country', '')}".strip(", ")
                
                entry = {
                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "company": pickup_data["CompanyName"],
                    "address": full_addr,
                    "date": pickup_data["PickupDate"],
                    "status": "Success" if prn else "Failed",
                    "prn": prn,
                    "details": pickup_data # Save full data
                }
                self.save_to_history(entry)

            except Exception as e:
                self.batch_tree.item(iid, values=(f"{pickup_data['Street']}, {pickup_data['City']}", "Error", str(e)))

        self.run_batch_btn.config(state=tk.NORMAL)
        self.stop_batch_btn.config(state=tk.DISABLED)
        self.next_btn.config(state=tk.DISABLED)
        messagebox.showinfo("Batch Complete", "Batch processing finished.")

    def cancel_pickup_gui(self):
        prn = simpledialog.askstring("Cancel Pickup", "Enter the PRN to cancel:", parent=self.root)
        if prn:
            try:
                result = self.api_client.cancel_pickup(prn)
                messagebox.showinfo("Cancel Result", json.dumps(result, indent=2))
            except Exception as e:
                messagebox.showerror("Cancel Error", str(e))

    def clear_fields(self):
        self.address_text.delete("1.0", tk.END)
        for entry in self.fields.values():
            entry.delete(0, tk.END)
        self.fields["Country"].insert(0, "CA")
        self.fields["CompanyName"].insert(0, "Omnitrans Inc")
        self.fields["ContactName"].insert(0, "Warehouse")

    def save_to_history(self, entry):
        history = []
        try:
            with open("pickup_history.json", "r") as f:
                history = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            history = []
        history.append(entry)
        with open(self.history_file, "w") as f:
            json.dump(history, f, indent=2)

    def show_history_window(self):
        top = tk.Toplevel(self.root)
        top.title("Pickup History")
        top.geometry("1000x500")
        
        # Search Frame
        search_frame = ttk.Frame(top, padding="10")
        search_frame.pack(fill=tk.X)
        ttk.Label(search_frame, text="Search (PRN/Address/Company):").pack(side=tk.LEFT)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=40)
        search_entry.pack(side=tk.LEFT, padx=10)
        search_entry.focus_set()

        search_entry.pack(side=tk.LEFT, padx=10)
        search_entry.focus_set()

        history = []
        try:
            with open(self.history_file, "r") as f:
                history = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            pass
            
        columns = ("Timestamp", "Company", "Address", "Date", "Time", "Status", "PRN")
        tree_frame = ttk.Frame(top)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def refresh_tree(*args):
            query = search_var.get().lower()
            for item in tree.get_children():
                tree.delete(item)
            for entry in history:
                match = any(query in str(entry.get(k, "")).lower() for k in ["prn", "address", "company"])
                if not query or match:
                    details = entry.get("details", {})
                    if isinstance(details, str):
                        try: details = json.loads(details)
                        except: details = {}
                    
                    ready_t = details.get("ReadyTime", "")
                    close_t = details.get("CloseTime", "")
                    time_str = f"{ready_t}-{close_t}" if ready_t else ""

                    tree.insert("", tk.END, values=(
                        entry.get("timestamp", ""),
                        entry.get("company", ""),
                        entry.get("address", ""),
                        entry.get("date", ""),
                        time_str,
                        entry.get("status", ""),
                        entry.get("prn", "")
                    ))

        search_var.trace_add("write", refresh_tree)
        
        for col in columns:
            tree.heading(col, text=col, command=lambda _col=col: self.sort_treeview(tree, _col, False))
            tree.column(col, width=120)
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Export Button
        export_btn = ttk.Button(search_frame, text="Export Selected to Excel", 
                               command=lambda: self.export_to_excel(tree, history))
        export_btn.pack(side=tk.LEFT, padx=10)
        
        def on_double_click(event):
            item_iid = tree.selection()[0]
            values = tree.item(item_iid, "values")
            prn = values[5]
            
            if prn:
                # PRN Verification Pop-up
                try:
                    res = self.api_client.get_pickup_status(prn)
                    # Pretty print the result in a new window
                    details_win = tk.Toplevel(top)
                    details_win.title(f"Pickup Details: {prn}")
                    details_win.geometry("500x400")
                    
                    # Extract useful fields from UPS response if possible
                    txt = tk.Text(details_win, padding=10, font=("Helvetica", 10))
                    txt.pack(fill=tk.BOTH, expand=True)
                    
                    # Simple formatting of the JSON response
                    txt.insert(tk.END, f"Live Pickup Information from UPS for PRN: {prn}\n\n")
                    if "PickupStatusResponse" in res:
                        status_data = res["PickupStatusResponse"]
                        txt.insert(tk.END, f"Status: {status_data.get('StatusCode', 'N/A')}\n")
                        txt.insert(tk.END, f"Pickup Date: {status_data.get('PickupDate', 'N/A')}\n")
                        txt.insert(tk.END, f"Ready Time: {status_data.get('ReadyTime', 'N/A')}\n")
                        txt.insert(tk.END, f"Close Time: {status_data.get('CloseTime', 'N/A')}\n\n")
                        
                    txt.insert(tk.END, "Full API Response:\n")
                    txt.insert(tk.END, json.dumps(res, indent=2))
                    txt.config(state=tk.DISABLED) # Read only
                except Exception as e:
                    messagebox.showerror("Error", f"Could not fetch pickup status: {str(e)}")
            else:
                # Fallback to Tracking search if no PRN
                entry = next((e for e in history if f"{e.get('Street', '')}, {e.get('City', '')}" == values[2]), None)
                if entry and entry.get("TrackingNumber"):
                    webbrowser.open(f"https://www.ups.com/track?tracknum={entry['TrackingNumber']}")

        tree.bind("<Double-1>", on_double_click)
        
        # UI Hint for Range Selection
        ttk.Label(top, text="Tip: Use Shift+Click to select a range of items, or Ctrl+Click to pick specific ones.", 
                  font=("Helvetica", 9, "italic"), foreground="gray").pack(side=tk.BOTTOM, pady=5)

        # Initial load
        refresh_tree()

    def sort_treeview(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    def export_to_excel(self, tree, history):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Export", "Please select one or more items to export.")
            return
            
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "UPS Pickups"
        
        headers = ["Created", "Company", "Full Address", "Date", "Ready", "Close", "Status", "PRN", "Tracking #"]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            
        for row_idx, item_iid in enumerate(selected, 2):
            values = tree.item(item_iid, "values")
            # values: (Timestamp, Company, Address, Date, Status, PRN)
            
            # More robust lookup: Match by timestamp and address
            entry = next((e for e in history if e.get("timestamp") == values[0] and e.get("address") == values[2]), None)
            
            if entry:
                details = entry.get("details", {})
                # Handle cases where details might be a JSON string or not a dict
                if isinstance(details, str):
                    try:
                        details = json.loads(details)
                    except:
                        details = {}
                if not isinstance(details, dict):
                    details = {}
                
                # Full Address: Rebuild with all components for completeness
                street = details.get("Street", "")
                city = details.get("City", "")
                state = details.get("State", "")
                zip_code = details.get("Zip", "")
                country = details.get("Country", "")
                
                if street:
                    full_address = f"{street}, {city}, {state} {zip_code}, {country}".strip(", ")
                else:
                    # Fallback to the top-level address if details are missing
                    full_address = entry.get("address", "")

                row_data = [
                    entry.get("timestamp"),
                    entry.get("company"),
                    full_address,
                    entry.get("date"),
                    details.get("ReadyTime"),
                    details.get("CloseTime"),
                    entry.get("status"),
                    entry.get("prn"),
                    details.get("TrackingNumber")
                ]
            else:
                # Fallback: Use values directly from Treeview if history match fails
                row_data = [
                    values[0], # Timestamp
                    values[1], # Company
                    values[2], # Full Address from Treeview
                    values[3], # Date
                    "", "", # Times placeholder
                    values[4], # Status
                    values[5], # PRN
                    "" # Tracking placeholder
                ]
                
            for col_idx, val in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=str(val) if val else "")
        
        filename = f"UPS_Pickups_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        try:
            wb.save(filename)
            messagebox.showinfo("Export Success", f"Exported {len(selected)} items to {filename}")
        except PermissionError:
            messagebox.showerror("Export Error", f"Could not save {filename}. Please close the file if it is open in Excel and try again.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to save file: {str(e)}")
            
        def cancel_selected():
            selected_items = tree.selection()
            if not selected_items: 
                messagebox.showwarning("Warning", "Please select at least one pickup to cancel.")
                return
            
            if not messagebox.askyesno("Confirm", f"Cancel {len(selected_items)} selected pickup(s)?"):
                return

            success_count = 0
            fail_count = 0
            
            for item_iid in selected_items:
                values = tree.item(item_iid, "values")
                prn = values[5] # PRN is the 6th column
                
                if prn:
                    # Find the entry in history
                    entry = next((e for e in history if e.get("prn") == prn), None)
                    if entry:
                        try:
                            res = self.api_client.cancel_pickup(prn)
                            if res.get("status") == "success" or "error" not in str(res).lower():
                                success_count += 1
                                tree.set(item_iid, "Status", "Cancelled")
                                entry["status"] = "Cancelled"
                            else:
                                fail_count += 1
                        except Exception as e:
                            fail_count += 1
                    else:
                        fail_count += 1
                else:
                    fail_count += 1

            # Update history file
            with open("pickup_history.json", "w") as f:
                json.dump(history, f, indent=2)

            messagebox.showinfo("Bulk Cancel Result", f"Finished.\nSuccess: {success_count}\nFailed: {fail_count}")

        def select_all():
            tree.selection_set(tree.get_children())

        def cancel_all_btn():
            select_all()
            cancel_selected()

        def clear_failed_cancelled():
            nonlocal history
            if not messagebox.askyesno("Confirm", "Remove all records with Status 'Failed' or 'Cancelled'?"):
                return
            new_history = [e for e in history if e.get("status") not in ["Failed", "Cancelled"]]
            deleted_count = len(history) - len(new_history)
            history[:] = new_history # Update in place
            with open(self.history_file, "w") as f:
                json.dump(history, f, indent=2)
            refresh_tree()
            messagebox.showinfo("Success", f"Cleared {deleted_count} record(s).")

        btn_frame = ttk.Frame(top)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Select All", command=select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel All", command=cancel_all_btn).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel Selected Pickup(s)", command=cancel_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Clear Failed/Cancelled", command=clear_failed_cancelled).pack(side=tk.LEFT, padx=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = UPSPickupGUI(root)
    root.mainloop()
