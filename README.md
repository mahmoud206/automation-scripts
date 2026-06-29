# ⚙️ Automation Scripts

A personal collection of Python scripts built to automate repetitive tasks — each script targets a specific workflow and runs independently.

---

## 📦 Scripts

| Script | Description |
|--------|-------------|
| [`catalog_report_generator.py`](./catalog_report_generator.py) | Extracts SKUs from catalog layout Excel files and generates a formatted stock analysis report |
| [`Spike_alarm.py`](./Spike_alarm.py) | Send Daily Email To The Supply chain manager To early Detect Stockout Skus |
| [`Quick_repeat_view.py`](./Quick_repeat_view.py) | A quick overview of the last 3 years’ sales compared with current stock helps in repeat decision-making |
| [`Quick_repeat_view.py`](./Daily_Auto_stk_checker.py) | Quick Daily Stock Comparison Between Today's and Yesterday's Stock, with Report Generation |
> More scripts will be added over time.

---

## 🛠️ Requirements

Each script has its own dependencies. Check the script header or the comments at the top of each file for install instructions.

Generally:

```bash
pip install pandas openpyxl
```

---

## 🚀 Usage

Each script is standalone — just run it directly:

```bash
python script_name.py
```

Scripts that need user input are fully interactive via the terminal.

---

## 📁 Structure

```
Automation/
├── README.md
│    └── catalog_report_generator.py
│    └── Spike_alarm.py
│    └── Quick repeat view.py


```

---
