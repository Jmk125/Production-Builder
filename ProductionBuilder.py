import tkinter as tk
from tkinter import ttk

class LaborSimulatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Construction Labor Simulator")

        self.workers = []
        self.tasks = []
        self.impact_entries = []

        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self.root, text="Simulation Name:").grid(row=0, column=0, sticky="w")
        self.sim_name_entry = ttk.Entry(self.root, width=30)
        self.sim_name_entry.grid(row=0, column=1, pady=5, sticky="w")

        self.restart_button = ttk.Button(self.root, text="🔄 Restart", command=self.restart)
        self.restart_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")

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

    def add_worker(self):
        row = len(self.workers)
        name_var = tk.StringVar()
        efficiency_var = tk.DoubleVar(value=1.0)

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

    def add_task(self):
        row = len(self.tasks)
        task_row_frame = ttk.Frame(self.task_frame)
        task_row_frame.grid(row=row, column=0, pady=4, sticky="w")

        task_name_var = tk.StringVar()
        base_time_var = tk.DoubleVar()
        time_unit_var = tk.StringVar(value="Minutes")
        material_unit_var = tk.StringVar(value="unit")
        assigned_workers = []

        ttk.Label(task_row_frame, text=f"Task {row+1}:").grid(row=0, column=0, padx=5, sticky="w")
        task_name_entry = ttk.Entry(task_row_frame, textvariable=task_name_var, width=20)
        task_name_entry.grid(row=0, column=1, padx=5)

        for i, worker in enumerate(self.workers):
            var = tk.BooleanVar()
            cb = ttk.Checkbutton(task_row_frame, text=worker["name_var"].get() or f"Worker {i+1}", variable=var)
            cb.grid(row=0, column=2+i, padx=3)
            assigned_workers.append(var)

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

    def setup_output_section(self):
        ttk.Label(self.output_frame, text="Output Type:").grid(row=0, column=0, sticky="w")
        self.output_type_var = tk.StringVar(value="Square-foot")
        output_dropdown = ttk.Combobox(self.output_frame, textvariable=self.output_type_var,
                                       values=["Square-foot", "Man Day"], state="readonly")
        output_dropdown.grid(row=0, column=1, sticky="w")
        output_dropdown.bind("<<ComboboxSelected>>", lambda e: self.update_output())

        ttk.Label(self.output_frame, text="Material Unit:").grid(row=1, column=0, sticky="e")
        self.material_unit_display = ttk.Entry(self.output_frame, state="readonly", width=15)
        self.material_unit_display.grid(row=1, column=1, padx=5, sticky="w")

        ttk.Label(self.output_frame, text="Length (ft):").grid(row=1, column=2, sticky="e")
        self.length_var = tk.DoubleVar(value=0)
        ttk.Entry(self.output_frame, textvariable=self.length_var, width=7).grid(row=1, column=3)

        ttk.Label(self.output_frame, text="Height (ft):").grid(row=1, column=4, sticky="e")
        self.height_var = tk.DoubleVar(value=0)
        ttk.Entry(self.output_frame, textvariable=self.height_var, width=7).grid(row=1, column=5)

        self.target_label = ttk.Label(self.output_frame, text="Target Area (sqft):")
        self.target_label.grid(row=2, column=0, sticky="e")
        self.target_area_var = tk.DoubleVar(value=0)
        self.target_entry = ttk.Entry(self.output_frame, textvariable=self.target_area_var, width=10)
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
        self.target_area_var.trace_add("write", lambda *args: self.update_output())
        self.time_display_unit.trace_add("write", lambda *args: self.update_output())
        self.output_type_var.trace_add("write", lambda *args: self.toggle_target_input())

    def toggle_target_input(self):
        mode = self.output_type_var.get()
        if mode == "Man Day":
            self.target_label.grid_remove()
            self.target_entry.grid_remove()
        else:
            self.target_label.grid()
            self.target_entry.grid()

    def add_impact(self):
        row = len(self.impact_entries) + 1
        name_var = tk.StringVar()
        time_var = tk.DoubleVar()
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

            unit_sqft = self.length_var.get() * self.height_var.get()
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
                total_area = self.target_area_var.get()
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

            elif mode == "Man Day":
                total_available_time = 8 * 60  # 8 hours in minutes
                effective_time = total_available_time - impact_time
                units_completed = effective_time / total_time_per_unit if total_time_per_unit > 0 else 0
                sqft_completed = units_completed * unit_sqft

                breakdown += f"\nAvailable time: 8 hrs = {total_available_time} min"
                breakdown += f"\n- Impacts: {impact_time:.2f} min → Working time = {effective_time:.2f} min"
                breakdown += f"\nTime per unit: {total_time_per_unit:.2f} min"
                breakdown += f"\nUnits completed: {effective_time:.2f} ÷ {total_time_per_unit:.2f} = {units_completed:.2f}"
                breakdown += f"\nUnit size: {unit_sqft:.2f} sqft → Total: {sqft_completed:.2f} sqft"

                self.final_output_label.config(
                    text=f"Total Production: {sqft_completed:.2f} sqft installed in 1 Man Day"
                )

            self.breakdown_label.config(text=breakdown)

        except Exception as e:
            self.final_output_label.config(text="Total Time: ???")
            self.breakdown_label.config(text="Calculation Error")

    def restart(self):
        self.workers.clear()
        self.tasks.clear()
        self.impact_entries.clear()
        for widget in self.worker_frame.winfo_children():
            widget.destroy()
        for widget in self.task_frame.winfo_children():
            widget.destroy()
        for widget in self.impact_frame.winfo_children():
            widget.destroy()
        self.sim_name_entry.delete(0, tk.END)
        self.length_var.set(0)
        self.height_var.set(0)
        self.target_area_var.set(0)
        self.time_display_unit.set("minutes")
        self.output_type_var.set("Square-foot")
        self.material_unit_display.config(state="normal")
        self.material_unit_display.delete(0, tk.END)
        self.material_unit_display.config(state="readonly")
        self.final_output_label.config(text="Total Time: ???")
        self.breakdown_label.config(text="")
        self.update_output()

def main():
    root = tk.Tk()
    app = LaborSimulatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
