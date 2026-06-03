import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import pandas as pd
import numpy as np
from scipy.stats import norm
import matplotlib.pyplot as plt

class GradingSystemApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Grading System Dashboard")
        self.root.geometry("550x650")
        
        # --- UI/UX Color Palette & Styling ---
        self.bg_color = "#f8f9fa"        # Light gray background
        self.primary_color = "#0056b3"   # Deep blue for headers/accents
        self.button_color = "#0d6efd"    # Bootstrap primary blue
        self.text_color = "#212529"      # Dark gray for text
        
        self.root.configure(bg=self.bg_color)
        
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure custom styles
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabel", background=self.bg_color, foreground=self.text_color, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=self.bg_color, foreground=self.primary_color, font=("Segoe UI", 14, "bold"))
        style.configure("TButton", background=self.button_color, foreground="white", font=("Segoe UI", 10, "bold"), padding=6)
        style.map("TButton", background=[("active", "#0b5ed7")])
        style.configure("TRadiobutton", background=self.bg_color, foreground=self.text_color, font=("Segoe UI", 10))

        # --- Data Attributes ---
        self.data = None
        self.default_stdev_adjustment = 1.0

        # Updated pre-adjustment boundaries for absolute grading
        self.default_boundaries = {
            90: "A", 85: "A-", 80: "B+", 75: "B", 70: "B-",
            65: "C+", 60: "C", 55: "C-", 50: "D", 0: "F"
        }

        # --- Tkinter Variables for Single-Page Input ---
        self.grading_scheme = tk.StringVar(value="absolute")
        
        # Dictionaries to store the entry variables for the unified form
        self.marks_vars = {comp: tk.StringVar() for comp in ["quizzes", "assignments", "midterm", "final", "project"]}
        self.weight_vars = {comp: tk.StringVar() for comp in ["quizzes", "assignments", "midterm", "final", "project"]}

        self.setup_ui()

    def setup_ui(self):
        """Builds the single-page front-end dashboard."""
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Header & File Upload
        ttk.Label(main_frame, text="1. Data Source", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        upload_frame = ttk.Frame(main_frame)
        upload_frame.pack(fill=tk.X, pady=(0, 20))
        self.upload_button = ttk.Button(upload_frame, text="Upload Excel File", command=self.upload_file)
        self.upload_button.pack(side=tk.LEFT)
        
        self.file_label = ttk.Label(upload_frame, text="No file selected...", foreground="#6c757d")
        self.file_label.pack(side=tk.LEFT, padx=15)

        # 2. Grading Scheme
        ttk.Label(main_frame, text="2. Grading Scheme", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))
        scheme_frame = ttk.Frame(main_frame)
        scheme_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Radiobutton(scheme_frame, text="Absolute", variable=self.grading_scheme, value="absolute").pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(scheme_frame, text="Relative", variable=self.grading_scheme, value="relative").pack(side=tk.LEFT)

        # 3. Component Configuration (Unified Form)
        ttk.Label(main_frame, text="3. Component Configuration", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        grid_frame = ttk.Frame(main_frame)
        grid_frame.pack(fill=tk.X, pady=(0, 30))

        # Grid Headers
        ttk.Label(grid_frame, text="Component", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Label(grid_frame, text="Total Marks", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, padx=20, pady=5)
        ttk.Label(grid_frame, text="Weightage (%)", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, pady=5)

        components = [
            ("Quizzes", "quizzes"), ("Assignments", "assignments"),
            ("Midterm", "midterm"), ("Finals", "final"), ("Project", "project")
        ]

        # Generate Entry Rows
        for i, (label, key) in enumerate(components, start=1):
            ttk.Label(grid_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=8)
            ttk.Entry(grid_frame, textvariable=self.marks_vars[key], width=12, justify="center").grid(row=i, column=1, padx=20, pady=8)
            ttk.Entry(grid_frame, textvariable=self.weight_vars[key], width=12, justify="center").grid(row=i, column=2, pady=8)

        # Horizontal Divider
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)

        # 4. Action Buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=10)
        
        self.process_button = ttk.Button(action_frame, text="Process Grades", command=self.process_grades)
        self.process_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))
        
        self.save_button = ttk.Button(action_frame, text="Save Results", command=self.save_results)
        self.save_button.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(10, 0))

    # --- Core Logic Methods (Unchanged working logic) ---

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                self.data = pd.read_excel(file_path, engine="openpyxl")
                required_columns = ['Quizzes', 'Assignments', 'Midterm', 'Finals', 'Project']
                if not all(col in self.data.columns for col in required_columns):
                    messagebox.showerror(
                        "Error",
                        "Excel file must contain required columns: Quizzes, Assignments, Midterm, Finals, Project.",
                    )
                    self.data = None
                    self.file_label.config(text="Invalid file structure.")
                else:
                    filename = file_path.split("/")[-1]
                    self.file_label.config(text=f"Loaded: {filename}", foreground="#198754") # Green text for success
                    messagebox.showinfo("Success", "File uploaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")

    def fetch_form_data(self):
        """Helper method to replace the old popups by reading from the new unified Tkinter variables."""
        weightages = {}
        total_marks = {}
        try:
            for key in self.weight_vars.keys():
                w_val = self.weight_vars[key].get()
                m_val = self.marks_vars[key].get()
                
                if not w_val or not m_val:
                    raise ValueError("All Total Marks and Weightage fields must be filled.")
                
                weightages[key] = float(w_val)
                total_marks[key] = float(m_val)

            if sum(weightages.values()) != 100:
                raise ValueError("Total weightages must equal exactly 100%.")
            
            return total_marks, weightages
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
            return None, None

    def calculate_final_scores(self, total_marks, weightages):
        try:
            self.data['Final Score'] = (
                (self.data['Quizzes'] / total_marks['quizzes']) * weightages['quizzes'] +
                (self.data['Assignments'] / total_marks['assignments']) * weightages['assignments'] +
                (self.data['Midterm'] / total_marks['midterm']) * weightages['midterm'] +
                (self.data['Finals'] / total_marks['final']) * weightages['final'] +
                (self.data['Project'] / total_marks['project']) * weightages['project']
            )
        except KeyError as e:
            messagebox.showerror("Error", f"Missing column in input data: {e}")

    def apply_absolute_grading(self, scores, boundaries):
        grades = []
        for score in scores:
            for boundary in sorted(boundaries.keys(), reverse=True):
                if score >= boundary:
                    grades.append(boundaries[boundary])
                    break
        return grades

    def apply_relative_grading(self, scores, stdev_adjustment):
        mean = np.mean(scores)
        std_dev = np.std(scores)
        grades = []

        for score in scores:
            if score >= mean + 2 * std_dev:
                grades.append("A")
            elif mean + (3 / 2) * std_dev <= score < mean + 2 * std_dev:
                grades.append("A-")
            elif mean + std_dev <= score < mean + (3 / 2) * std_dev:
                grades.append("B+")
            elif mean + (1 / 2) * std_dev <= score < mean + std_dev:
                grades.append("B")
            elif mean - (1 / 2) * std_dev <= score < mean + (1 / 2) * std_dev:
                grades.append("B-")
            elif mean - std_dev <= score < mean - (1 / 2) * std_dev:
                grades.append("C+")
            elif mean - (4 / 3) * std_dev <= score < mean - std_dev:
                grades.append("C")
            elif mean - (5 / 3) * std_dev <= score < mean - (4 / 3) * std_dev:
                grades.append("C-")
            elif mean - 2 * std_dev <= score < mean - (5 / 3) * std_dev:
                grades.append("D")
            else:
                grades.append("F")
        return grades

    def apply_post_adjustment_grading(self, scores, stdev_multiplier):
        mean = np.mean(scores)
        std_dev = np.std(scores)
        grades = []

        grade_boundaries = [
            ("A", mean + 1.5 * std_dev),
            ("A-", mean + (1.5 * std_dev - stdev_multiplier * std_dev)),
            ("B+", mean + (1.5 * std_dev - 2 * stdev_multiplier * std_dev)),
            ("B", mean + (1.5 * std_dev - 3 * stdev_multiplier * std_dev)),
            ("B-", mean + (1.5 * std_dev - 4 * stdev_multiplier * std_dev)),
            ("C+", mean + (1.5 * std_dev - 5 * stdev_multiplier * std_dev)),
            ("C", mean + (1.5 * std_dev - 6 * stdev_multiplier * std_dev)),
            ("C-", mean + (1.5 * std_dev - 7 * stdev_multiplier * std_dev)),
            ("D", mean + (1.5 * std_dev - 8 * stdev_multiplier * std_dev)),
            ("F", float("-inf"))
        ]

        for score in scores:
            for grade, boundary in grade_boundaries:
                if score >= boundary:
                    grades.append(grade)
                    break

        return grades

    def visualize_grades(self, pre_grades, post_grades, title):
        pre_grade_counts = pd.Series(pre_grades).value_counts().sort_index()
        post_grade_counts = pd.Series(post_grades).value_counts().sort_index()

        fig, axes = plt.subplots(1, 2, figsize=(12, 6))
        pre_grade_counts.plot(kind="bar", alpha=0.7, color="skyblue", ax=axes[0])
        post_grade_counts.plot(kind="bar", alpha=0.7, color="lightgreen", ax=axes[1])

        axes[0].set_title("Pre-Adjustment Grades")
        axes[0].set_xlabel("Grades")
        axes[0].set_ylabel("Frequency")
        axes[1].set_title("Post-Adjustment Grades")
        axes[1].set_xlabel("Grades")
        axes[1].set_ylabel("Frequency")

        plt.suptitle(title)
        plt.tight_layout()
        plt.show()

    def visualize_bell_curve(self, scores, title):
        mean = np.mean(scores)
        std_dev = np.std(scores)

        x = np.linspace(mean - 3 * std_dev, mean + 3 * std_dev, 1000)
        y = norm.pdf(x, mean, std_dev)

        plt.figure(figsize=(10, 6))
        plt.plot(x, y, label="Bell Curve", color="red", linewidth=2)
        plt.xlabel("Grades")
        plt.ylabel("Probability Density")
        plt.title(f"{title}\nMean: {mean:.2f}, Std Dev: {std_dev:.2f}")
        plt.axvline(mean, color="blue", linestyle="--", label="Mean")
        plt.axvline(mean + std_dev, color="green", linestyle="--", label="Mean + 1 SD")
        plt.axvline(mean - std_dev, color="green", linestyle="--", label="Mean - 1 SD")
        plt.axvline(mean + 2 * std_dev, color="orange", linestyle="--", label="Mean + 2 SD")
        plt.axvline(mean - 2 * std_dev, color="orange", linestyle="--", label="Mean - 2 SD")
        plt.axvline(mean + 3 * std_dev, color="purple", linestyle="--", label="Mean + 3 SD")
        plt.axvline(mean - 3 * std_dev, color="purple", linestyle="--", label="Mean - 3 SD")
        plt.legend()
        plt.show()

    def input_absolute_boundaries(self):
        boundaries = {}
        try:
            prev_boundary = float('inf')
            while True:
                grade = simpledialog.askstring("Input", "Enter grade label (e.g., A, B, etc.): ")
                if not grade:
                    break
                boundary_input = simpledialog.askstring("Input", f"Enter lower boundary for grade {grade}: ")
                if not boundary_input:
                    break
                try:
                    boundary = float(boundary_input)
                    if boundary >= prev_boundary:
                        messagebox.showerror("Input Error", f"The lower boundary for {grade} must be less than the previous boundary.")
                        continue
                    boundaries[boundary] = grade
                    prev_boundary = boundary
                except ValueError:
                    messagebox.showerror("Input Error", "Invalid input for boundary. Please enter a numeric value.")
                    continue

                more = messagebox.askyesno("Continue", "Do you want to add another grade boundary?")
                if not more:
                    break

            boundaries = dict(sorted(boundaries.items(), reverse=True))
        except ValueError:
            messagebox.showerror("Input Error", "Invalid input for grade boundaries.")
            return None
        return boundaries

    def process_grades(self):
        if self.data is None:
            messagebox.showerror("Error", "Please upload a valid grades file first.")
            return

        # Fetch the inputs from the new unified form
        total_marks, weightages = self.fetch_form_data()
        if not total_marks or not weightages:
            return

        # Calculate the final scores using the unified inputs
        self.calculate_final_scores(total_marks, weightages)

        scheme = self.grading_scheme.get()
        if scheme == "absolute":
            pre_grades = self.apply_absolute_grading(self.data['Final Score'], self.default_boundaries)
            self.visualize_grades(pre_grades, pre_grades, "Pre-Adjustment Absolute Grading")

            adjust = messagebox.askyesno("Adjust Grades", "Do you want to adjust the grade boundaries?")
            if adjust:
                post_boundaries = self.input_absolute_boundaries()
                if post_boundaries:
                    post_grades = self.apply_absolute_grading(self.data['Final Score'], post_boundaries)
                    self.visualize_grades(pre_grades, post_grades, "Post-Adjustment Absolute Grading")
                    self.data['Grade'] = post_grades
                    self.show_summary_statistics(pre_grades, post_grades)
                else:
                    messagebox.showerror("Error", "Invalid grade boundaries entered.")
                    return
            else:
                self.data['Grade'] = pre_grades

        elif scheme == "relative":
            pre_grades = self.apply_relative_grading(self.data['Final Score'], self.default_stdev_adjustment)
            self.visualize_grades(pre_grades, pre_grades, "Pre-Adjustment Relative Grading")
            self.visualize_bell_curve(self.data['Final Score'], "Pre-Adjustment Relative Grading Bell Curve")

            adjust = messagebox.askyesno("Adjust Grades", "Do you want to adjust the grading using a custom standard deviation multiplier?")
            if adjust:
                multiplier_str = simpledialog.askstring("Input", "Enter standard deviation multiplier for post-adjustment (default is 1.0): ")
                if multiplier_str:
                    try:
                        stdev_multiplier = float(multiplier_str)
                        post_grades = self.apply_post_adjustment_grading(self.data['Final Score'], stdev_multiplier)
                        self.visualize_grades(pre_grades, post_grades, "Post-Adjustment Relative Grading")
                        self.visualize_bell_curve(self.data['Final Score'], "Post-Adjustment Relative Grading Bell Curve")
                        self.data['Grade'] = post_grades
                        self.show_summary_statistics(pre_grades, post_grades)
                    except ValueError:
                        messagebox.showerror("Input Error", "Invalid multiplier. Processing halted.")
                        return
            else:
                self.data['Grade'] = pre_grades

        messagebox.showinfo("Success", "Grades processed successfully!")

    def show_summary_statistics(self, pre_grades, post_grades):
        pre_grade_counts = pd.Series(pre_grades).value_counts()
        post_grade_counts = pd.Series(post_grades).value_counts()

        summary = "Summary Statistics:\n"
        summary += f"Total Students: {len(self.data)}\n"
        summary += "Grade Changes:\n"

        for grade in pre_grade_counts.index:
            pre_count = pre_grade_counts.get(grade, 0)
            post_count = post_grade_counts.get(grade, 0)
            if pre_count != post_count:
                summary += f"{grade}: {pre_count} -> {post_count}\n"

        messagebox.showinfo("Summary Statistics", summary)

    def save_results(self):
        if self.data is None or 'Grade' not in self.data.columns:
            messagebox.showerror("Error", "No data to save or grades not assigned. Please process grades first.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", ".csv"), ("Excel Files", ".xlsx")])
        if file_path:
            try:
                output_data = self.data.copy()

                if 'Final Score' not in output_data.columns:
                    output_data['Final Score'] = self.data['Final Score']
                if 'Grade' not in output_data.columns:
                    output_data['Grade'] = self.data['Grade']

                if file_path.endswith('.csv'):
                    output_data.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    output_data.to_excel(file_path, index=False, engine='openpyxl')

                messagebox.showinfo("Success", f"Results saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GradingSystemApp(root)
    root.mainloop()