import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pandas as pd
import os
from datetime import datetime

class LaborSimulatorApp:
    VERSION = "V1.0"
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"Construction Labor Simulator {self.VERSION}")

        self.workers = []
        self.tasks = []
        self.impact_entries = []
        self.current_file_path = None

        self.create_widgets()

    def create_widgets(self):
        # Top frame for simulation name and control buttons
        top_frame = ttk.Frame(self.root)
        top_frame.grid(row=0, column=0, columnspan=4, sticky="ew", padx=10, pady=5)
        
        ttk.Label(top_frame, text="Simulation Name:").grid(row=0, column=0, sticky="w")
        self.sim_name_entry = ttk.Entry(top_frame, width=30)
        self.sim_name_entry.grid(row=0, column=1, pady=5, sticky="w")

        # Buttons frame
        button_frame = ttk.Frame(top_frame)
        button_frame.grid(row=0, column=2, padx=5, pady=5, sticky="e")
        
        self.save_button = ttk.Button(button_frame, text="💾 Save", command=self.save_simulation)
        self.save_button.pack(side=tk.LEFT, padx=2)
        
        self.load_button = ttk.Button(button_frame, text="📂 Load", command=self.load_simulation)
        self.load_button.pack(side=tk.LEFT, padx=2)
        
        self.export_button = ttk.Button(button_frame, text="📊 Export", command=self.export_to_excel)
        self.export_button.pack(side=tk.LEFT, padx=2)
        
        self.restart_button = ttk.Button(button_frame, text="🔄 Restart", command=self.restart)
        self.restart_button.pack(side=tk.LEFT, padx=2)

        # Continue with your existing widget creation
        ttk.Label(self.root, text="Crew:").grid(row=1, column=0, sticky="w")
        self.add_worker_button = ttk.Button(self.root, text="➕ Add Worker", command=self.add_worker)
        self.add_worker_button.grid(row=1, column=1, sticky="w")

        self.worker_frame = ttk.LabelFrame(self.root, text="Workers")
        self.worker_frame.grid(row=2, column=0, columnspan=4, sticky="we", padx=10, pady=5)

        ttk.Label(self.root, text="Tasks:").grid(row=3, column=0, sticky="w")
        self.add_task_button = ttk.Button(self.root, text="➕ Add Task", command=self.add_task)
        self.add_task_button.grid(row=3, column=1, sticky="w")

        self.task_frame = ttk.LabelFrame(self.root, text="Tasks")
        self.task_frame.grid(row=4, column=0, columnspan=4, sticky="we", padx=10, pady=5)

        self.output_frame = ttk.LabelFrame(self.root, text="Output")
        self.output_frame.grid(row=5, column=0, columnspan=4, sticky="we", padx=10, pady=10)

        self.setup_output_section()
        
        # Add version label at the bottom
        version_frame = ttk.Frame(self.root)
        version_frame.grid(row=6, column=0, columnspan=4, sticky="ew", padx=10, pady=(5, 10))
        version_label = ttk.Label(
            version_frame, 
            text=f"Construction Labor Simulator {self.VERSION} | © 2025 Jmk125",
            foreground="gray"
        )
        version_label.pack(side=tk.RIGHT)

    def setup_output_section(self):
        ttk.Label(self.output_frame, text="Output Type:").grid(row=0, column=0, sticky="w")
        self.output_type_var = tk.StringVar(value="Square-foot")
        output_dropdown = ttk.Combobox(self.output_frame, textvariable=self.output_type_var,
                                       values=["Square-foot", "Linear-Foot", "Man Day (SF)", "Man Day (LF)"], state="readonly")
        output_dropdown.grid(row=0, column=1, sticky="w")
        output_dropdown.bind("<<ComboboxSelected>>", lambda e: self.update_output_mode())

        ttk.Label(self.output_frame, text="Material Unit:").grid(row=1, column=0, sticky="e")
        self.material_unit_display = ttk.Entry(self.output_frame, state="readonly", width=15)
        self.material_unit_display.grid(row=1, column=1, padx=5, sticky="w")

        # Length and height row (will show/hide based on mode)
        self.dimensions_frame = ttk.Frame(self.output_frame)
        self.dimensions_frame.grid(row=1, column=2, columnspan=4, sticky="w")

        ttk.Label(self.dimensions_frame, text="Length (ft):").grid(row=0, column=0, sticky="e")
        self.length_var = tk.DoubleVar(value=0)
        ttk.Entry(self.dimensions_frame, textvariable=self.length_var, width=7).grid(row=0, column=1)

        # Height label and entry (will be hidden for Linear-Foot mode)
        self.height_label = ttk.Label(self.dimensions_frame, text="Height (ft):")
        self.height_label.grid(row=0, column=2, sticky="e")
        self.height_var = tk.DoubleVar(value=0)
        self.height_entry = ttk.Entry(self.dimensions_frame, textvariable=self.height_var, width=7)
        self.height_entry.grid(row=0, column=3)

        # Target area/length label and entry
        self.target_label = ttk.Label(self.output_frame, text="Target Area (sqft):")
        self.target_label.grid(row=2, column=0, sticky="e")
        self.target_var = tk.DoubleVar(value=0)
        self.target_entry = ttk.Entry(self.output_frame, textvariable=self.target_var, width=10)
        self.target_entry.grid(row=2, column=1)

        ttk.Label(self.output_frame, text="Display Time In:").grid(row=2, column=2, sticky="e")
        self.time_display_unit = tk.StringVar(value="minutes")
        ttk.Combobox(self.output_frame, textvariable=self.time_display_unit,
                     values=["minutes", "hours"], state="readonly", width=10).grid(row=2, column=3)

        self.impact_frame = ttk.Frame(self.output_frame)
        self.impact_frame.grid(row=3, column=0, columnspan=6, pady=(10, 0), sticky="w")
        ttk.Button(self.impact_frame, text="➕ Add Impact", command=self.add_impact).grid(row=0, column=0)

        self.final_output_label = ttk.Label(self.output_frame, text="Total Time: ???")
        self.final_output_label.grid(row=4, column=0, columnspan=6, pady=10, sticky="w")

        self.breakdown_label = ttk.Label(self.output_frame, text="", justify="left")
        self.breakdown_label.grid(row=5, column=0, columnspan=6, sticky="w")

        self.length_var.trace_add("write", lambda *args: self.update_output())
        self.height_var.trace_add("write", lambda *args: self.update_output())
        self.target_var.trace_add("write", lambda *args: self.update_output())
        self.time_display_unit.trace_add("write", lambda *args: self.update_output())
        
        # Initial UI setup based on default mode
        self.update_output_mode()

    def update_output_mode(self):
        """Update UI based on the selected output mode."""
        mode = self.output_type_var.get()
        
        # Configure target field label and UI elements based on mode
        if mode == "Square-foot":
            self.target_label.config(text="Target Area (sqft):")
            self.target_entry.grid()
            self.height_label.grid()
            self.height_entry.grid()
        elif mode == "Linear-Foot":
            self.target_label.config(text="Target Length (lf):")
            self.target_entry.grid()
            self.height_label.grid_remove()
            self.height_entry.grid_remove()
        elif mode == "Man Day (SF)":
            self.target_label.grid_remove()
            self.target_entry.grid_remove()
            self.height_label.grid()
            self.height_entry.grid()
        elif mode == "Man Day (LF)":
            self.target_label.grid_remove()
            self.target_entry.grid_remove()
            self.height_label.grid_remove()
            self.height_entry.grid_remove()
        
        # Update calculations
        self.update_output()

    def toggle_target_input(self):
        """
        Legacy method kept for compatibility.
        Use update_output_mode() instead for new code.
        """
        self.update_output_mode()

    def save_simulation(self):
        """Save the current simulation to a JSON file."""
        # Collect all simulation data
        simulation_data = {
            "simulation_name": self.sim_name_entry.get(),
            "workers": self.collect_worker_data(),
            "tasks": self.collect_task_data(),
            "impacts": self.collect_impact_data(),
            "output_settings": {
                "output_type": self.output_type_var.get(),
                "length": self.length_var.get(),
                "height": self.height_var.get(),
                "target": self.target_var.get(),
                "time_display_unit": self.time_display_unit.get()
            },
            "version": self.VERSION
        }
        
        # Ask for a save location if not already saving to a file
        file_path = self.current_file_path
        if not file_path:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialfile=f"{self.sim_name_entry.get() or 'simulation'}.json"
            )
            if not file_path:  # User cancelled
                return
            self.current_file_path = file_path
            
        try:
            with open(file_path, 'w') as f:
                json.dump(simulation_data, f, indent=4)
            messagebox.showinfo("Success", f"Simulation saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save simulation: {str(e)}")

    def collect_worker_data(self):
        """Collect data from all workers."""
        worker_data = []
        for worker in self.workers:
            worker_data.append({
                "name": worker["name_var"].get(),
                "efficiency": worker["efficiency_var"].get()
            })
        return worker_data
    
    def collect_task_data(self):
        """Collect data from all tasks."""
        task_data = []
        for task in self.tasks:
            # Get assigned workers
            assigned = []
            for i, var in enumerate(task["assigned_workers"]):
                if var.get():
                    assigned.append(i)
                    
            task_data.append({
                "name": task["task_name_var"].get(),
                "base_time": task["base_time_var"].get(),
                "time_unit": task["time_unit_var"].get(),
                "material_unit": task["material_unit_var"].get(),
                "assigned_workers": assigned
            })
        return task_data
    
    def collect_impact_data(self):
        """Collect data from all impact entries."""
        impact_data = []
        for name_var, time_var in self.impact_entries:
            impact_data.append({
                "name": name_var.get(),
                "time": time_var.get()
            })
        return impact_data

    def load_simulation(self):
        """Load a simulation from a JSON file."""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not file_path:  # User cancelled
            return
            
        try:
            with open(file_path, 'r') as f:
                simulation_data = json.load(f)
                
            # Reset current state before loading
            self.restart()
            
            # Set the current file path
            self.current_file_path = file_path
            
            # Load simulation name
            self.sim_name_entry.delete(0, tk.END)
            self.sim_name_entry.insert(0, simulation_data.get("simulation_name", ""))
            
            # Load workers
            for worker_data in simulation_data.get("workers", []):
                self.add_worker(worker_data)
                
            # Load tasks
            for task_data in simulation_data.get("tasks", []):
                self.add_task(task_data)
                
            # Load impacts
            for impact_data in simulation_data.get("impacts", []):
                self.add_impact(impact_data)
                
            # Load output settings
            output_settings = simulation_data.get("output_settings", {})
            
            # Handle backward compatibility for renamed "Man Day" option
            output_type = output_settings.get("output_type", "Square-foot")
            if output_type == "Man Day":
                output_type = "Man Day (SF)"
            self.output_type_var.set(output_type)
            
            self.length_var.set(output_settings.get("length", 0))
            self.height_var.set(output_settings.get("height", 0))
            
            # Support for both old and new formats
            if "target_area" in output_settings:
                self.target_var.set(output_settings.get("target_area", 0))
            else:
                self.target_var.set(output_settings.get("target", 0))
                
            self.time_display_unit.set(output_settings.get("time_display_unit", "minutes"))
            
            # Check if the version is compatible
            file_version = simulation_data.get("version", "V1.0")
            if file_version != self.VERSION:
                messagebox.showwarning("Version Mismatch", 
                    f"The file was created with version {file_version}, current version is {self.VERSION}. Some features may not work correctly.")
            
            # Update UI based on loaded data
            self.update_output_mode()
            messagebox.showinfo("Success", f"Simulation loaded from {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load simulation: {str(e)}")

    def add_worker(self, worker_data=None):
        """Add a new worker, optionally with pre-loaded data."""
        row = len(self.workers)
        name_var = tk.StringVar()
        efficiency_var = tk.DoubleVar(value=1.0)
        
        # If loading from a file, set the values
        if worker_data:
            name_var.set(worker_data.get("name", ""))
            efficiency_var.set(worker_data.get("efficiency", 1.0))

        name_entry = ttk.Entry(self.worker_frame, textvariable=name_var, width=20)
        efficiency_entry = ttk.Entry(self.worker_frame, textvariable=efficiency_var, width=10)

        ttk.Label(self.worker_frame, text=f"Worker {row+1}").grid(row=row, column=0, padx=5, pady=2, sticky="e")
        name_entry.grid(row=row, column=1, padx=5, pady=2)
        ttk.Label(self.worker_frame, text="Efficiency:").grid(row=row, column=2, sticky="e")
        efficiency_entry.grid(row=row, column=3, padx=5, pady=2)

        self.workers.append({
            "name_var": name_var,
            "efficiency_var": efficiency_var
        })

        # Update existing tasks with the new worker checkboxes
        for task in self.tasks:
            task_row_frame = task["task_row_frame"]
            var = tk.BooleanVar()
            cb = ttk.Checkbutton(task_row_frame, text=name_var.get() or f"Worker {len(self.workers)}", variable=var)
            cb.grid(row=0, column=2+len(task["assigned_workers"]), padx=3)
            task["assigned_workers"].append(var)
            var.trace_add("write", task["update_adjusted_time"])

    def add_task(self, task_data=None):
        """Add a new task, optionally with pre-loaded data."""
        row = len(self.tasks)
        task_row_frame = ttk.Frame(self.task_frame)
        task_row_frame.grid(row=row, column=0, pady=4, sticky="w")

        task_name_var = tk.StringVar()
        base_time_var = tk.DoubleVar()
        time_unit_var = tk.StringVar(value="Minutes")
        material_unit_var = tk.StringVar(value="unit")
        assigned_workers = []

        # If loading from a file, set the values
        if task_data:
            task_name_var.set(task_data.get("name", ""))
            base_time_var.set(task_data.get("base_time", 0.0))
            time_unit_var.set(task_data.get("time_unit", "Minutes"))
            material_unit_var.set(task_data.get("material_unit", "unit"))

        ttk.Label(task_row_frame, text=f"Task {row+1}:").grid(row=0, column=0, padx=5, sticky="w")
        task_name_entry = ttk.Entry(task_row_frame, textvariable=task_name_var, width=20)
        task_name_entry.grid(row=0, column=1, padx=5)

        for i, worker in enumerate(self.workers):
            var = tk.BooleanVar()
            cb = ttk.Checkbutton(task_row_frame, text=worker["name_var"].get() or f"Worker {i+1}", variable=var)
            cb.grid(row=0, column=2+i, padx=3)
            assigned_workers.append(var)
            
            # If loading from a file, set worker assignments
            if task_data and "assigned_workers" in task_data and i in task_data["assigned_workers"]:
                var.set(True)

        ttk.Label(task_row_frame, text="Base Time:").grid(row=1, column=0, sticky="e", padx=5)
        base_time_entry = ttk.Entry(task_row_frame, textvariable=base_time_var, width=10)
        base_time_entry.grid(row=1, column=1, padx=5, sticky="w")

        ttk.Label(task_row_frame, text="Time Unit:").grid(row=1, column=2, sticky="e")
        time_unit_menu = ttk.Combobox(task_row_frame, textvariable=time_unit_var,
                                      values=["Minutes", "Hours"], state="readonly", width=10)
        time_unit_menu.grid(row=1, column=3, padx=5, sticky="w")

        ttk.Label(task_row_frame, text="Per:").grid(row=1, column=4, sticky="e")
        material_unit_entry = ttk.Entry(task_row_frame, textvariable=material_unit_var, width=10)
        material_unit_entry.grid(row=1, column=5, padx=5, sticky="w")

        result_label = ttk.Label(task_row_frame, text="Adjusted Time: ???")
        result_label.grid(row=1, column=6, padx=10, sticky="w")

        def update_adjusted_time(*args):
            selected_efficiencies = []
            worker_count = 0
            for i, var in enumerate(assigned_workers):
                if var.get():
                    try:
                        eff = float(self.workers[i]["efficiency_var"].get())
                        selected_efficiencies.append(eff)
                        worker_count += 1
                    except:
                        continue

            if selected_efficiencies and worker_count > 0:
                avg_eff = sum(selected_efficiencies) / len(selected_efficiencies)
                try:
                    base_time = base_time_var.get()
                    time_unit = time_unit_var.get().lower()
                    unit_factor = 1 if time_unit == "minutes" else 60
                    adjusted_time = (base_time * unit_factor / avg_eff) * worker_count
                    material_unit = material_unit_var.get()
                    result_label.config(text=f"Adjusted Time: {adjusted_time:.2f} minutes/{material_unit}")
                except:
                    result_label.config(text="Adjusted Time: ???")
            else:
                result_label.config(text="Adjusted Time: ???")

            self.update_output()

        base_time_var.trace_add("write", update_adjusted_time)
        time_unit_var.trace_add("write", update_adjusted_time)
        material_unit_var.trace_add("write", update_adjusted_time)
        for worker_var in assigned_workers:
            worker_var.trace_add("write", update_adjusted_time)
        for worker in self.workers:
            worker["efficiency_var"].trace_add("write", update_adjusted_time)

        self.tasks.append({
            "task_row_frame": task_row_frame,
            "task_name_var": task_name_var,
            "assigned_workers": assigned_workers,
            "base_time_var": base_time_var,
            "time_unit_var": time_unit_var,
            "material_unit_var": material_unit_var,
            "result_label": result_label,
            "update_adjusted_time": update_adjusted_time
        })
        
        # Trigger update to recalculate times
        update_adjusted_time()

    def add_impact(self, impact_data=None):
        """Add a new impact, optionally with pre-loaded data."""
        row = len(self.impact_entries) + 1
        name_var = tk.StringVar()
        time_var = tk.DoubleVar()
        
        # If loading from a file, set the values
        if impact_data:
            name_var.set(impact_data.get("name", ""))
            time_var.set(impact_data.get("time", 0.0))

        ttk.Entry(self.impact_frame, textvariable=name_var, width=20).grid(row=row, column=0, padx=5, pady=2)
        ttk.Entry(self.impact_frame, textvariable=time_var, width=10).grid(row=row, column=1, padx=5, pady=2)
        ttk.Label(self.impact_frame, text="minutes").grid(row=row, column=2)
        time_var.trace_add("write", lambda *args: self.update_output())
        self.impact_entries.append((name_var, time_var))
        self.update_output()

    def update_output(self):
        if not self.tasks:
            return

        try:
            breakdown = ""
            material_unit = self.tasks[0]["material_unit_var"].get()
            self.material_unit_display.config(state="normal")
            self.material_unit_display.delete(0, tk.END)
            self.material_unit_display.insert(0, material_unit)
            self.material_unit_display.config(state="readonly")

            unit_length = self.length_var.get()
            unit_height = self.height_var.get()
            unit_sqft = unit_length * unit_height
            total_workers = len(self.workers)

            total_time_per_unit = 0
            for task in self.tasks:
                selected_efficiencies = []
                worker_count = 0
                for i, assigned in enumerate(task["assigned_workers"]):
                    if assigned.get():
                        eff = float(self.workers[i]["efficiency_var"].get())
                        selected_efficiencies.append(eff)
                        worker_count += 1

                if selected_efficiencies and worker_count > 0:
                    avg_eff = sum(selected_efficiencies) / len(selected_efficiencies)
                    base_time = task["base_time_var"].get()
                    time_unit = task["time_unit_var"].get().lower()
                    unit_factor = 1 if time_unit == "minutes" else 60
                    adjusted_time = (base_time * unit_factor / avg_eff) * worker_count
                    total_time_per_unit += adjusted_time
                    breakdown += f"- {task['task_name_var'].get() or 'Task'}: {adjusted_time:.2f} min/unit\n"

            impact_time = sum(t[1].get() * total_workers for t in self.impact_entries)
            for name_var, time_var in self.impact_entries:
                breakdown += f"- Impact '{name_var.get() or 'Unnamed'}': {time_var.get():.2f} min × {total_workers} workers = {time_var.get() * total_workers:.2f} min\n"

            mode = self.output_type_var.get()
            if mode == "Square-foot":
                total_area = self.target_var.get()
                units_needed = total_area / unit_sqft if unit_sqft > 0 else 0
                total_time = total_time_per_unit * units_needed + impact_time
                breakdown += f"\nUnits needed: {units_needed:.2f} → Task time: {total_time_per_unit:.2f} × {units_needed:.2f} = {total_time_per_unit * units_needed:.2f} min"
                breakdown += f"\n+ Impacts: {impact_time:.2f} min → Total: {total_time:.2f} min"

                if self.time_display_unit.get() == "hours":
                    total_time /= 60
                    breakdown += f"\n\nDisplayed in hours: {total_time:.2f} hours"

                self.final_output_label.config(
                    text=f"Total Time: {total_time:.2f} {self.time_display_unit.get()} to complete {total_area:.0f} sqft"
                )
                
            elif mode == "Linear-Foot":
                target_length = self.target_var.get()
                # For linear foot, units needed is directly target_length / unit_length
                units_needed = target_length / unit_length if unit_length > 0 else 0
                total_time = total_time_per_unit * units_needed + impact_time
                
                breakdown += f"\nTarget length: {target_length:.2f} lf"
                breakdown += f"\nUnit length: {unit_length:.2f} lf"
                breakdown += f"\nUnits needed: {target_length:.2f} ÷ {unit_length:.2f} = {units_needed:.2f} units"
                breakdown += f"\nTask time: {total_time_per_unit:.2f} min/unit × {units_needed:.2f} units = {total_time_per_unit * units_needed:.2f} min"
                breakdown += f"\n+ Impacts: {impact_time:.2f} min → Total: {total_time:.2f} min"

                if self.time_display_unit.get() == "hours":
                    total_time /= 60
                    breakdown += f"\n\nDisplayed in hours: {total_time:.2f} hours"

                self.final_output_label.config(
                    text=f"Total Time: {total_time:.2f} {self.time_display_unit.get()} to complete {target_length:.0f} lf"
                )

            elif mode == "Man Day (SF)":
                # Calculate man-day as 8 hours per worker
                hours_per_worker = 8
                total_available_time = hours_per_worker * 60 * total_workers  # Convert to minutes
                effective_time = total_available_time - impact_time
                units_completed = effective_time / total_time_per_unit if total_time_per_unit > 0 else 0
                sqft_completed = units_completed * unit_sqft

                breakdown += f"\nCrew size: {total_workers} workers × 8 hours = {total_workers * 8} work hours"
                breakdown += f"\nTotal available time: {total_available_time} minutes"
                breakdown += f"\n- Impacts: {impact_time:.2f} min → Working time = {effective_time:.2f} min"
                breakdown += f"\nTime per unit: {total_time_per_unit:.2f} min"
                breakdown += f"\nUnits completed: {effective_time:.2f} ÷ {total_time_per_unit:.2f} = {units_completed:.2f}"
                breakdown += f"\nUnit size: {unit_sqft:.2f} sqft → Total: {sqft_completed:.2f} sqft"

                self.final_output_label.config(
                    text=f"Total Production: {sqft_completed:.2f} sqft installed in 1 Man Day ({total_workers} workers)"
                )
                
            elif mode == "Man Day (LF)":
                # Calculate man-day as 8 hours per worker
                hours_per_worker = 8
                total_available_time = hours_per_worker * 60 * total_workers  # Convert to minutes
                effective_time = total_available_time - impact_time
                units_completed = effective_time / total_time_per_unit if total_time_per_unit > 0 else 0
                lf_completed = units_completed * unit_length  # Linear feet completed

                breakdown += f"\nCrew size: {total_workers} workers × 8 hours = {total_workers * 8} work hours"
                breakdown += f"\nTotal available time: {total_available_time} minutes"
                breakdown += f"\n- Impacts: {impact_time:.2f} min → Working time = {effective_time:.2f} min"
                breakdown += f"\nTime per unit: {total_time_per_unit:.2f} min"
                breakdown += f"\nUnits completed: {effective_time:.2f} ÷ {total_time_per_unit:.2f} = {units_completed:.2f}"
                breakdown += f"\nUnit length: {unit_length:.2f} lf → Total: {lf_completed:.2f} linear feet"

                self.final_output_label.config(
                    text=f"Total Production: {lf_completed:.2f} linear feet installed in 1 Man Day ({total_workers} workers)"
                )

            self.breakdown_label.config(text=breakdown)

        except Exception as e:
            self.final_output_label.config(text="Total Time: ???")
            self.breakdown_label.config(text="Calculation Error")

    def export_to_excel(self):
        """Export the simulation data to an Excel file."""
        if not self.tasks:
            messagebox.showwarning("Warning", "No tasks to export.")
            return
            
        try:
            # Ask for save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"{self.sim_name_entry.get() or 'simulation'}.xlsx"
            )
            if not file_path:  # User cancelled
                return
                
            # Create a dict to hold all dataframes for Excel
            excel_data = {}
            
            # Workers data
            worker_data = []
            for i, worker in enumerate(self.workers):
                worker_data.append({
                    "Worker #": i+1,
                    "Name": worker["name_var"].get(),
                    "Efficiency": worker["efficiency_var"].get()
                })
            if worker_data:
                excel_data["Workers"] = pd.DataFrame(worker_data)
            
            # Tasks data
            task_data = []
            for i, task in enumerate(self.tasks):
                # Get assigned workers
                assigned_workers = []
                for j, var in enumerate(task["assigned_workers"]):
                    if var.get():
                        worker_name = self.workers[j]["name_var"].get() or f"Worker {j+1}"
                        assigned_workers.append(worker_name)
                
                task_data.append({
                    "Task #": i+1,
                    "Name": task["task_name_var"].get(),
                    "Base Time": task["base_time_var"].get(),
                    "Time Unit": task["time_unit_var"].get(),
                    "Material Unit": task["material_unit_var"].get(),
                    "Assigned Workers": ", ".join(assigned_workers)
                })
            if task_data:
                excel_data["Tasks"] = pd.DataFrame(task_data)
            
            # Impact data
            impact_data = []
            for i, (name_var, time_var) in enumerate(self.impact_entries):
                impact_data.append({
                    "Impact #": i+1,
                    "Name": name_var.get(),
                    "Time (min)": time_var.get()
                })
            if impact_data:
                excel_data["Impacts"] = pd.DataFrame(impact_data)
            
            # Simulation settings
            mode = self.output_type_var.get()
            current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            settings_data = [{
                "Simulation Name": self.sim_name_entry.get(),
                "Output Type": mode,
                "Length (ft)": self.length_var.get(),
                "Time Display Unit": self.time_display_unit.get(),
                "Export Date": current_date,
                "User": "Jmk125",  # Hardcoded username
                "Version": self.VERSION
            }]
            
            # Add height for Square-foot and Man Day (SF) modes
            if mode in ["Square-foot", "Man Day (SF)"]:
                settings_data[0]["Height (ft)"] = self.height_var.get()
                
            # Add target field with appropriate label
            if mode == "Square-foot":
                settings_data[0]["Target Area (sqft)"] = self.target_var.get()
            elif mode == "Linear-Foot":
                settings_data[0]["Target Length (lf)"] = self.target_var.get()
                
            excel_data["Settings"] = pd.DataFrame(settings_data)
            
            # Create a worksheet with calculation results
            results_data = []
            
            material_unit = self.tasks[0]["material_unit_var"].get() if self.tasks else "unit"
            unit_length = self.length_var.get()
            unit_height = self.height_var.get()
            unit_sqft = unit_length * unit_height
            total_workers = len(self.workers)
            total_time_per_unit = 0
            
            # Task time calculations
            for task in self.tasks:
                selected_efficiencies = []
                worker_count = 0
                for i, assigned in enumerate(task["assigned_workers"]):
                    if assigned.get():
                        eff = float(self.workers[i]["efficiency_var"].get())
                        selected_efficiencies.append(eff)
                        worker_count += 1

                if selected_efficiencies and worker_count > 0:
                    avg_eff = sum(selected_efficiencies) / len(selected_efficiencies)
                    base_time = task["base_time_var"].get()
                    time_unit = task["time_unit_var"].get().lower()
                    unit_factor = 1 if time_unit == "minutes" else 60
                    adjusted_time = (base_time * unit_factor / avg_eff) * worker_count
                    total_time_per_unit += adjusted_time
                    
                    results_data.append({
                        "Category": "Task",
                        "Name": task["task_name_var"].get() or f"Task {len(results_data)+1}",
                        "Time (min)": adjusted_time,
                        "Notes": f"Base: {base_time} {time_unit}, Efficiency: {avg_eff:.2f}, Workers: {worker_count}"
                    })
            
            # Impact times
            impact_time = sum(t[1].get() * total_workers for t in self.impact_entries)
            for name_var, time_var in self.impact_entries:
                results_data.append({
                    "Category": "Impact",
                    "Name": name_var.get() or "Unnamed Impact",
                    "Time (min)": time_var.get() * total_workers,
                    "Notes": f"Per Worker: {time_var.get()} min × {total_workers} workers"
                })
            
            # Final calculation based on output type
            if mode == "Square-foot":
                total_area = self.target_var.get()
                units_needed = total_area / unit_sqft if unit_sqft > 0 else 0
                total_time = total_time_per_unit * units_needed + impact_time
                
                results_data.append({
                    "Category": "Summary",
                    "Name": "Units Needed",
                    "Time (min)": None,
                    "Notes": f"{units_needed:.2f} units for {total_area:.2f} sqft"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Task Time",
                    "Time (min)": total_time_per_unit * units_needed,
                    "Notes": f"{total_time_per_unit:.2f} min/unit × {units_needed:.2f} units"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Impact Time",
                    "Time (min)": impact_time,
                    "Notes": "Sum of all impacts"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Time",
                    "Time (min)": total_time,
                    "Notes": f"To complete {total_area:.2f} sqft"
                })
                if self.time_display_unit.get() == "hours":
                    results_data.append({
                        "Category": "Summary",
                        "Name": "Total Time (hours)",
                        "Time (min)": total_time / 60,
                        "Notes": f"Converted to hours: {total_time/60:.2f}"
                    })
            elif mode == "Linear-Foot":
                target_length = self.target_var.get()
                units_needed = target_length / unit_length if unit_length > 0 else 0
                total_time = total_time_per_unit * units_needed + impact_time
                
                results_data.append({
                    "Category": "Summary",
                    "Name": "Target Length",
                    "Time (min)": None,
                    "Notes": f"{target_length:.2f} linear feet"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Unit Length",
                    "Time (min)": None,
                    "Notes": f"{unit_length:.2f} feet per unit"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Units Needed",
                    "Time (min)": None,
                    "Notes": f"{units_needed:.2f} units ({target_length:.2f} lf ÷ {unit_length:.2f} lf/unit)"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Task Time",
                    "Time (min)": total_time_per_unit * units_needed,
                    "Notes": f"{total_time_per_unit:.2f} min/unit × {units_needed:.2f} units"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Impact Time",
                    "Time (min)": impact_time,
                    "Notes": "Sum of all impacts"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Time",
                    "Time (min)": total_time,
                    "Notes": f"To complete {target_length:.2f} linear feet"
                })
                if self.time_display_unit.get() == "hours":
                    results_data.append({
                        "Category": "Summary",
                        "Name": "Total Time (hours)",
                        "Time (min)": total_time / 60,
                        "Notes": f"Converted to hours: {total_time/60:.2f}"
                    })
            elif mode == "Man Day (SF)":
                # Calculate man-day as 8 hours per worker
                hours_per_worker = 8
                total_available_time = hours_per_worker * 60 * total_workers  # Minutes
                effective_time = total_available_time - impact_time
                units_completed = effective_time / total_time_per_unit if total_time_per_unit > 0 else 0
                sqft_completed = units_completed * unit_sqft
                
                results_data.append({
                    "Category": "Summary",
                    "Name": "Crew Size",
                    "Time (min)": None,
                    "Notes": f"{total_workers} workers × 8 hours = {total_workers * 8} work hours"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Available Time",
                    "Time (min)": total_available_time,
                    "Notes": f"{total_workers} workers × 8 hours = {total_available_time} minutes"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Effective Working Time",
                    "Time (min)": effective_time,
                    "Notes": f"Available time minus impacts: {total_available_time} - {impact_time:.2f}"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Units Completed",
                    "Time (min)": None,
                    "Notes": f"{units_completed:.2f} units ({effective_time:.2f} min ÷ {total_time_per_unit:.2f} min/unit)"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Production (SF)",
                    "Time (min)": None,
                    "Notes": f"{sqft_completed:.2f} sqft ({units_completed:.2f} units × {unit_sqft:.2f} sqft/unit)"
                })
            elif mode == "Man Day (LF)":
                # Calculate man-day as 8 hours per worker
                hours_per_worker = 8
                total_available_time = hours_per_worker * 60 * total_workers  # Minutes
                effective_time = total_available_time - impact_time
                units_completed = effective_time / total_time_per_unit if total_time_per_unit > 0 else 0
                lf_completed = units_completed * unit_length
                
                results_data.append({
                    "Category": "Summary",
                    "Name": "Crew Size",
                    "Time (min)": None,
                    "Notes": f"{total_workers} workers × 8 hours = {total_workers * 8} work hours"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Available Time",
                    "Time (min)": total_available_time,
                    "Notes": f"{total_workers} workers × 8 hours = {total_available_time} minutes"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Effective Working Time",
                    "Time (min)": effective_time,
                    "Notes": f"Available time minus impacts: {total_available_time} - {impact_time:.2f}"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Units Completed",
                    "Time (min)": None,
                    "Notes": f"{units_completed:.2f} units ({effective_time:.2f} min ÷ {total_time_per_unit:.2f} min/unit)"
                })
                results_data.append({
                    "Category": "Summary",
                    "Name": "Total Production (LF)",
                    "Time (min)": None,
                    "Notes": f"{lf_completed:.2f} linear feet ({units_completed:.2f} units × {unit_length:.2f} lf/unit)"
                })
            
            excel_data["Results"] = pd.DataFrame(results_data)
            
            # Write to Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for sheet_name, df in excel_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
            messagebox.showinfo("Success", f"Data exported to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")

    def restart(self):
        self.workers.clear()
        self.tasks.clear()
        self.impact_entries.clear()
        self.current_file_path = None
        for widget in self.worker_frame.winfo_children():
            widget.destroy()
        for widget in self.task_frame.winfo_children():
            widget.destroy()
        for widget in self.impact_frame.winfo_children():
            widget.destroy()
        self.sim_name_entry.delete(0, tk.END)
        self.length_var.set(0)
        self.height_var.set(0)
        self.target_var.set(0)
        self.time_display_unit.set("minutes")
        self.output_type_var.set("Square-foot")
        self.material_unit_display.config(state="normal")
        self.material_unit_display.delete(0, tk.END)
        self.material_unit_display.config(state="readonly")
        self.final_output_label.config(text="Total Time: ???")
        self.breakdown_label.config(text="")
        ttk.Button(self.impact_frame, text="➕ Add Impact", command=self.add_impact).grid(row=0, column=0)
        self.update_output()

def main():
    root = tk.Tk()
    app = LaborSimulatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
