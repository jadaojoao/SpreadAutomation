import customtkinter as ctk
import traceback
import threading
from app.widgets import FilePickerWidget, ProgressWidget
from engine.workflow_1a import Mode1AWorkflow
from core.models import EntityType

class Screen1A(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        
        self.title = ctk.CTkLabel(self, text="Mode 1A: Fill Existing Spread", font=ctk.CTkFont(size=20, weight="bold"))
        self.title.grid(row=0, column=0, pady=10, padx=10, sticky="w")
        
        # Files selection
        self.source_picker = FilePickerWidget(self, "Source (CVM Excel):")
        self.source_picker.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        
        self.spread_picker = FilePickerWidget(self, "Spread Proxy:")
        self.spread_picker.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        
        # Settings
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        
        # Company Code
        ctk.CTkLabel(self.settings_frame, text="Company:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.company_var = ctk.StringVar()
        ctk.CTkEntry(self.settings_frame, textvariable=self.company_var, width=120).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Period
        ctk.CTkLabel(self.settings_frame, text="Period (e.g. 4T24):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.period_var = ctk.StringVar()
        ctk.CTkEntry(self.settings_frame, textvariable=self.period_var, width=100).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # Entity
        self.entity_var = ctk.StringVar(value="consolidated")
        ctk.CTkRadioButton(self.settings_frame, text="Consolidated", variable=self.entity_var, value="consolidated").grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        ctk.CTkRadioButton(self.settings_frame, text="Individual", variable=self.entity_var, value="individual").grid(row=1, column=2, columnspan=2, padx=5, pady=5, sticky="w")

        # Column Mapping settings
        self.col_frame = ctk.CTkFrame(self)
        self.col_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)

        ctk.CTkLabel(self.col_frame, text="Dest Column:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.dest_col_var = ctk.StringVar(value="C")
        ctk.CTkEntry(self.col_frame, textvariable=self.dest_col_var, width=60).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(self.col_frame, text="Prior Col (Opt):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.prior_col_var = ctk.StringVar(value="D")
        ctk.CTkEntry(self.col_frame, textvariable=self.prior_col_var, width=60).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # Actions
        self.run_btn = ctk.CTkButton(self, text="Execute Mode 1A", command=self.run_workflow)
        self.run_btn.grid(row=5, column=0, pady=15, padx=10)
        
        # Progress & Status
        self.progress = ProgressWidget(self)
        self.progress.grid(row=6, column=0, sticky="ew", padx=10, pady=5)
        
        # Logs
        self.log_box = ctk.CTkTextbox(self, height=120)
        self.log_box.grid(row=7, column=0, sticky="nsew", padx=10, pady=(5, 10))
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(7, weight=1)

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    def run_workflow(self):
        source = self.source_picker.get_path()
        spread = self.spread_picker.get_path()
        period = self.period_var.get()
        company = self.company_var.get()
        entity = self.entity_var.get()
        dest_col = self.dest_col_var.get()
        prior_col = self.prior_col_var.get()
        
        if not all([source, spread, period, company, dest_col]):
            self.log("[!] Error: You need to specify source, spread, company, period, and destination column.")
            return

        self.run_btn.configure(state="disabled")
        self.progress.update_progress("Starting...", 0.1)
        self.log(f"[*] Dispatching Mode 1A for {company} {period}...")

        threading.Thread(target=self._execute, args=(source, spread, company, period, dest_col, prior_col, entity), daemon=True).start()

    def _execute(self, source, spread, company, period, dest_col, prior_col, entity):
        try:
            entity_enum = EntityType.CONSOLIDATED if entity == "consolidated" else EntityType.INDIVIDUAL
            workflow = Mode1AWorkflow()
            
            self.progress.update_progress("Running workflow in engine...", 0.5)
            
            result = workflow.execute(
                source_path=source,
                spread_path=spread,
                company=company,
                period=period,
                dest_col=dest_col,
                prior_col=prior_col if getattr(prior_col, "strip", lambda:"" )() else None,
                entity_type=entity_enum
            )
            
            self.progress.update_progress("Done!", 1.0)
            self.log("[+] Workflow completed successfully!")
            
            val = result.get("validation", {})
            self.log(f"    Mapped: {result.get('mapped_count')} items.")
            self.log(f"    Validity: {val.get('is_valid')} (Missing: {len(val.get('missing', []))})")
            
        except Exception as e:
            self.progress.update_progress("Failed", 0.0)
            self.log(f"[-] Workflow failed:\n{traceback.format_exc()}")
        finally:
            self.run_btn.configure(state="normal")
