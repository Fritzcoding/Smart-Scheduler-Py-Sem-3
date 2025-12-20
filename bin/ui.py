"""
Schedule Creator v3.0 - Clean single-file UI implementation

Features:
- Activity manager dialog
- Group organizer dialog with auto-distribute
- Weekly and daily schedule generation
- JSON/CSV/Image export, basic analysis

This file is the clean, replacement implementation. To deploy it as the active UI,
rename/move it to `ui.py` (you may need to remove or back up the existing `ui.py`).
"""

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QListWidget, QListWidgetItem, QSpinBox, QComboBox
from PyQt5.QtCore import Qt
from word import make_word_doc, make_word_doc_only
import traceback
import time
import random
import json
import csv
import os
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw
import cv2


class ActivityManagerDialog(QDialog):
    """Dialog for managing custom activities"""
    
    def __init__(self, activities=None, parent=None):
        super().__init__(parent)
        self.activities = activities if activities else []
        self.setWindowTitle("Manage Activities")
        self.setGeometry(100, 100, 400, 300)
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout()
        
        title = QLabel("Custom Activities")
        title.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(title)
        
        input_layout = QHBoxLayout()
        self.activity_input = QLineEdit()
        self.activity_input.setPlaceholderText("Enter activity name")
        add_btn = QPushButton("Add Activity")
        add_btn.clicked.connect(self.add_activity)
        input_layout.addWidget(self.activity_input)
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)
        
        list_label = QLabel("Current Activities:")
        layout.addWidget(list_label)
        
        self.activity_list = QListWidget()
        for activity in self.activities:
            item = QListWidgetItem(activity)
            item.setFlags(item.flags() | Qt.ItemIsSelectable)
            self.activity_list.addItem(item)
        layout.addWidget(self.activity_list)
        
        remove_layout = QHBoxLayout()
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self.remove_activity)
        remove_layout.addWidget(remove_btn)
        remove_layout.addStretch()
        layout.addLayout(remove_layout)
        
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("OK")
        cancel_btn = QPushButton("Cancel")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def add_activity(self):
        text = self.activity_input.text().strip()
        if text and text not in self.activities:
            self.activities.append(text)
            item = QListWidgetItem(text)
            self.activity_list.addItem(item)
            self.activity_input.clear()
    
    def remove_activity(self):
        for item in self.activity_list.selectedItems():
            self.activities.remove(item.text())
            self.activity_list.takeItem(self.activity_list.row(item))
    
    def get_activities(self):
        return self.activities


class GroupOrganizerDialog(QDialog):
    """Dialog for managing groups and participants"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.groups = {}
        self.setWindowTitle("Organize Groups")
        self.setGeometry(100, 100, 500, 400)
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout()
        
        num_layout = QHBoxLayout()
        num_layout.addWidget(QLabel("Number of Groups:"))
        self.num_groups_spin = QSpinBox()
        self.num_groups_spin.setMinimum(2)
        self.num_groups_spin.setMaximum(10)
        self.num_groups_spin.setValue(4)
        self.num_groups_spin.valueChanged.connect(self.update_groups)
        num_layout.addWidget(self.num_groups_spin)
        num_layout.addStretch()
        layout.addLayout(num_layout)
        
        layout.addWidget(QLabel("Add Participants:"))
        input_layout = QHBoxLayout()
        self.participant_input = QLineEdit()
        self.participant_input.setPlaceholderText("Enter participant name")
        add_btn = QPushButton("Add Participant")
        add_btn.clicked.connect(self.add_participant)
        input_layout.addWidget(self.participant_input)
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)
        
        self.participant_list = QListWidget()
        layout.addWidget(self.participant_list)
        
        distribute_btn = QPushButton("Auto-Distribute to Groups")
        distribute_btn.clicked.connect(self.distribute_participants)
        layout.addWidget(distribute_btn)
        
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("OK")
        cancel_btn = QPushButton("Cancel")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def update_groups(self):
        num = self.num_groups_spin.value()
        if not self.groups or len(self.groups) != num:
            self.groups = {f"Group {i+1}": [] for i in range(num)}
    
    def add_participant(self):
        text = self.participant_input.text().strip()
        if text:
            self.participant_list.addItem(text)
            self.participant_input.clear()
    
    def distribute_participants(self):
        if not self.groups:
            self.update_groups()
        
        for group in self.groups:
            self.groups[group] = []
        
        participants = []
        for i in range(self.participant_list.count()):
            participants.append(self.participant_list.item(i).text())
        
        random.shuffle(participants)
        num_groups = len(self.groups)
        for idx, participant in enumerate(participants):
            group_idx = idx % num_groups
            group_name = list(self.groups.keys())[group_idx]
            self.groups[group_name].append(participant)
        
        msg = "Distribution Complete:\n"
        for group, members in self.groups.items():
            msg += f"{group}: {', '.join(members)}\n"
        
        QMessageBox.information(self, "Auto-Distribution", msg)
    
    def get_groups(self):
        return self.groups


class Ui_MainWindow(object):
    """Main UI class for Schedule Creator"""
    
    def setupUi(self, MainWindow):
        # keep a reference to the actual QMainWindow so dialogs can use it as parent
        self.main_window = MainWindow
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        MainWindow.setWindowTitle("Schedule Creator v3.0")
        
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        MainWindow.setCentralWidget(self.centralwidget)
        
        layout = QVBoxLayout()
        
        title = QLabel("Schedule Creator")
        title_font = QtGui.QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)
        
        activity_layout = QHBoxLayout()
        activity_layout.addWidget(QLabel("Activities:"))
        self.activity_manager_btn = QPushButton("Manage Activities")
        self.activity_manager_btn.clicked.connect(self.open_activity_manager)
        activity_layout.addWidget(self.activity_manager_btn)
        activity_layout.addStretch()
        layout.addLayout(activity_layout)
        
        # Group display panel
        group_display_layout = QHBoxLayout()
        group_display_layout.addWidget(QLabel("Groups Configured:"))
        self.group_count_label = QLabel("0")
        self.group_count_label.setStyleSheet("font-weight: bold; color: #2196F3;")
        group_display_layout.addWidget(self.group_count_label)
        self.group_members_label = QLabel("(no groups yet)")
        self.group_members_label.setWordWrap(True)
        group_display_layout.addWidget(self.group_members_label)
        group_display_layout.addStretch()
        layout.addLayout(group_display_layout)
        
        group_layout = QHBoxLayout()
        group_layout.addWidget(QLabel("Groups:"))
        self.group_organizer_btn = QPushButton("Organize Groups")
        self.group_organizer_btn.clicked.connect(self.open_group_organizer)
        group_layout.addWidget(self.group_organizer_btn)
        group_layout.addStretch()
        layout.addLayout(group_layout)
        
        period_layout = QHBoxLayout()
        period_layout.addWidget(QLabel("Periods per Day:"))
        self.periods_spin = QSpinBox()
        self.periods_spin.setMinimum(2)
        # enforce maximum of 6 periods per day
        self.periods_spin.setMaximum(6)
        self.periods_spin.setValue(6)
        self.periods_spin.setToolTip("Maximum 6 periods per day")
        period_layout.addWidget(self.periods_spin)
        period_layout.addStretch()
        layout.addLayout(period_layout)
        
        week_layout = QHBoxLayout()
        week_layout.addWidget(QLabel("Week:"))
        self.week_combo = QComboBox()
        for i in range(1, 8):
            self.week_combo.addItem(f"Week {i}")
        week_layout.addWidget(self.week_combo)
        week_layout.addStretch()
        layout.addLayout(week_layout)
        
        self.generate_btn = QPushButton("Generate Schedule")
        self.generate_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.generate_btn.clicked.connect(self.generate_schedule)
        layout.addWidget(self.generate_btn)
        
        export_layout = QHBoxLayout()
        export_layout.addWidget(QLabel("Export:"))
        
        json_btn = QPushButton("JSON")
        json_btn.clicked.connect(self.export_json)
        export_layout.addWidget(json_btn)
        
        csv_btn = QPushButton("CSV")
        csv_btn.clicked.connect(self.export_csv)
        export_layout.addWidget(csv_btn)

        image_btn = QPushButton("Image")
        image_btn.clicked.connect(self.export_image)
        export_layout.addWidget(image_btn)

        # Word-only export button (creates only the .docx)
        word_btn = QPushButton("Word")
        word_btn.clicked.connect(self.export_word)
        export_layout.addWidget(word_btn)

        # Pie chart export button
        pie_btn = QPushButton("Pie")
        pie_btn.clicked.connect(self.export_pie)
        export_layout.addWidget(pie_btn)

        analyze_btn = QPushButton("Analyze")
        analyze_btn.clicked.connect(self.analyze)
        export_layout.addWidget(analyze_btn)
        
        export_layout.addStretch()
        layout.addLayout(export_layout)
        
        layout.addStretch()
        self.centralwidget.setLayout(layout)
        
        # Initialize with good default example
        self.activities = ["Soccer", "Basketball", "Hockey", "Tennis", "Volleyball", "Badminton", "Ping Pong", "Gymnastics"]
        self.groups = {
            "Group 1": ["Alice", "Bob"],
            "Group 2": ["Charlie", "Diana"],
            "Group 3": ["Eve", "Frank"],
            "Group 4": ["Grace", "Henry"]
        }
        self.matrix = None
        self.start_time = None
        self.max_activity_uses = 2  # how many times each activity can be used per group
        self.update_group_display()
    
    def open_activity_manager(self):
        try:
            dialog = ActivityManagerDialog(self.activities, self.main_window)
            if dialog.exec_():
                self.activities = dialog.get_activities()
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            QMessageBox.critical(None, "Activity Manager Error", f"An error occurred opening Activity Manager:\n{str(e)}\n\nSee console for traceback.")
    
    def open_group_organizer(self):
        try:
            dialog = GroupOrganizerDialog(self.main_window)
            if dialog.exec_():
                self.groups = dialog.get_groups()
                self.update_group_display()
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            QMessageBox.critical(None, "Group Organizer Error", f"An error occurred opening Group Organizer:\n{str(e)}\n\nSee console for traceback.")
    
    def generate_schedule(self):
        if not self.activities:
            self.show_error("Please add activities first", "No Activities")
            return
        
        if not self.groups:
            self.show_error("Please organize groups first", "No Groups")
            return
        
        try:
            num_groups = len(self.groups)
            # validate basic constraints before attempting to solve
            issues = self.validate_constraints(num_groups)
            if issues:
                msg = "Cannot generate schedule because of the following constraint(s):\n\n"
                msg += "\n".join(f"- {it}" for it in issues)
                self.show_error(msg, "Constraint Violation")
                return

            # Only weekly mode supported
            self.matrix = self.generate_weekly_matrix(num_groups)

            self.start_time = time.time()
            if self.solve():
                make_word_doc(self.matrix, self.week_combo.currentText())
                try:
                    self.export_pie()
                except Exception:
                    pass
                self.show_info("Schedule generated successfully!", "Success")
            else:
                # attempt to auto-adjust max_activity_uses if capacity is the issue
                if self.try_auto_adjust_and_solve(num_groups):
                    make_word_doc(self.matrix, self.week_combo.currentText())
                    try:
                        self.export_pie()
                    except Exception:
                        pass
                    self.show_info(f"Schedule generated successfully!\n(Adjusted activity limit to {self.max_activity_uses})", "Success")
                else:
                    # provide more help when solver fails without obvious constraints
                    self.show_error("Could not generate schedule with these constraints.\nTry adding more activities, more participants, or reducing periods per day.", "Failed")
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            self.show_error(f"Error: {str(e)}\n\nSee console for traceback.", "Error")
    
    def generate_weekly_matrix(self, num_groups):
        matrix = []
        num_periods = self.periods_spin.value()
        num_days = 4
        for group_idx in range(num_groups):
            group_matrix = [[] for _ in range(num_periods)]
            for period in range(num_periods):
                for day in range(num_days):
                    group_matrix[period].append("")
            matrix.append(group_matrix)
        return matrix
    
    def generate_daily_matrix(self, num_groups):
        # Daily mode removed; keep the function but raise so it is not used.
        raise NotImplementedError("Daily schedule mode has been removed. Use Weekly mode.")
    
    def solve(self):
        if time.time() - self.start_time > 5:
            return False
        
        find = self.find_empty()
        if not find:
            return True
        
        group, period, day = find
        # randomize activity order to reduce repetitive first-period assignments
        choices = self.activities[:]
        random.shuffle(choices)
        for activity in choices:
            if self.valid(activity, (group, period, day)):
                self.matrix[group][period][day] = activity
                
                if self.solve():
                    return True
        
        self.matrix[group][period][day] = ""
        return False
    
    def find_empty(self):
        for g, group in enumerate(self.matrix):
            for p, period in enumerate(group):
                for d, day in enumerate(period):
                    if day == "":
                        return (g, p, d)
        return None
    
    def valid(self, activity, pos):
        g, p, d = pos
        
        for group_idx in range(len(self.matrix)):
            if group_idx != g and self.matrix[group_idx][p][d] == activity:
                return False
        
        count = sum(1 for period in self.matrix[g] for day in period if day == activity)
        if count >= self.max_activity_uses:
            return False
        
        return True
    
    def export_json(self):
        matrix = self._matrix_or_template()
        if matrix is None:
            return

        try:
            schedule_dict = {}
            for g, group in enumerate(matrix):
                schedule_dict[f"Group {g+1}"] = group

            filepath = f"Generated Schedules/{self.week_combo.currentText()}_schedule.json"
            os.makedirs("Generated Schedules", exist_ok=True)
            with open(filepath, 'w') as f:
                json.dump(schedule_dict, f, indent=2)

            self.show_info(f"Exported to {filepath}", "Export Complete")
        except Exception as e:
            self.show_error(f"Error: {str(e)}", "Export Failed")

    def export_word(self):
        matrix = self._matrix_or_template()
        if matrix is None:
            return

        try:
            make_word_doc_only(matrix, self.week_combo.currentText())
            filepath = f"Generated Schedules/{self.week_combo.currentText()} Schedules.docx"
            self.show_info(f"Exported Word document to {filepath}", "Word Export Complete")
        except Exception as e:
            self.show_error(f"Error: {str(e)}", "Word Export Failed")

    def export_pie(self):
        matrix = self._matrix_or_template()
        if matrix is None:
            return

        try:
            activity_counts = {}
            for group in matrix:
                for period in group:
                    for activity in period:
                        if activity and activity != "":
                            activity_counts[activity] = activity_counts.get(activity, 0) + 1

            # If no activities are assigned in the matrix, fall back to the
            # available `self.activities` so the user can export a pie chart
            # showing available activities rather than requiring a generated schedule.
            if not activity_counts:
                if getattr(self, 'activities', None):
                    activity_counts = {a: 1 for a in self.activities}
                else:
                    self.show_error("No activity data to plot", "Pie Export Failed")
                    return

            labels = list(activity_counts.keys())
            sizes = list(activity_counts.values())
            cmap = plt.get_cmap('tab20')
            colors = [cmap(i % 20) for i in range(len(labels))]

            fig, ax = plt.subplots(figsize=(8, 8))
            explode = [0.04] * len(labels)
            wedges, texts, autotexts = ax.pie(
                sizes,
                labels=labels,
                autopct='%1.1f%%',
                startangle=140,
                colors=colors,
                explode=explode,
                pctdistance=0.77,
                wedgeprops={'edgecolor': 'white', 'linewidth': 0.7}
            )
            ax.axis('equal')
            plt.setp(autotexts, size=10, weight='bold', color='white')
            plt.title(f'Activity Distribution - {self.week_combo.currentText()}')
            filepath = f"Generated Schedules/{self.week_combo.currentText()}_pie.png"
            os.makedirs("Generated Schedules", exist_ok=True)
            plt.tight_layout()
            plt.savefig(filepath, dpi=150)
            plt.close()

            self.show_info(f"Pie chart saved to {filepath}", "Pie Export Complete")
        except Exception as e:
            self.show_error(f"Error: {str(e)}", "Pie Export Failed")
    
    def export_csv(self):
        matrix = self._matrix_or_template()
        if matrix is None:
            return

        try:
            filepath = f"Generated Schedules/{self.week_combo.currentText()}_schedule.csv"
            os.makedirs("Generated Schedules", exist_ok=True)
            with open(filepath, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Group", "Period", "Days"])

                for g, group in enumerate(matrix):
                    for p, period in enumerate(group):
                        writer.writerow([f"Group {g+1}", f"Period {p+1}"] + period)

            self.show_info(f"Exported to {filepath}", "CSV Export Complete")
        except Exception as e:
            self.show_error(f"Error: {str(e)}", "CSV Export Failed")
    
    def export_image(self):
        matrix = self._matrix_or_template()
        if matrix is None:
            return

        try:
            img = Image.new('RGB', (1000, 600 + len(matrix) * 100), color='white')
            draw = ImageDraw.Draw(img)

            y_pos = 20
            for g, group in enumerate(matrix):
                text = f"Group {g+1}: {' | '.join([str(period) for period in group])}"
                draw.text((20, y_pos), text, fill='black')
                y_pos += 100

            filepath = f"Generated Schedules/{self.week_combo.currentText()}_schedule.png"
            os.makedirs("Generated Schedules", exist_ok=True)
            img.save(filepath)

            self.show_info(f"Image saved to {filepath}", "Image Export Complete")
        except Exception as e:
            self.show_error(f"Error: {str(e)}", "Image Export Failed")
    
    def analyze(self):
        matrix = self._matrix_or_template()
        if matrix is None:
            return

        try:
            activity_counts = {}
            for group in matrix:
                for period in group:
                    for activity in period:
                        if activity:
                            activity_counts[activity] = activity_counts.get(activity, 0) + 1
            
            counts = np.array(list(activity_counts.values()))
            total_unique = len(activity_counts)
            total_assign = int(np.sum(counts)) if counts.size > 0 else 0
            avg_freq = float(np.mean(counts)) if counts.size > 0 else 0.0
            std_dev = float(np.std(counts)) if counts.size > 0 else 0.0

            msg = (
                "Schedule Analysis:\n\n"
                "Total Unique Activities: {}\n"
                "Total Assignments: {}\n"
                "Average Activity Frequency: {:.2f}\n"
                "Std Deviation: {:.2f}".format(total_unique, total_assign, avg_freq, std_dev)
            )

            self.show_info(msg, "Schedule Analysis")
        except Exception as e:
            self.show_error(f"Error: {str(e)}", "Analysis Failed")

    def _matrix_or_template(self):
        """Return the current matrix or build a blank template from current groups/periods.

        Returns None and shows an error if no groups are defined.
        """
        if self.matrix:
            return self.matrix

        num_groups = len(self.groups)
        if num_groups == 0:
            self.show_error("No groups defined to create a schedule template.", "Export Failed")
            return None

        # build a blank weekly matrix using current periods setting
        return self.generate_weekly_matrix(num_groups)

    # Daily mode removed: no get_time_label or update_time_selector
    
    def try_auto_adjust_and_solve(self, num_groups):
        """
        Attempt to auto-adjust max_activity_uses to meet capacity requirements,
        then retry solving. Returns True if successful, False otherwise.
        """
        num_activities = len(self.activities)
        num_periods = self.periods_spin.value()
        slots_per_group = num_periods * 4
        
        # calculate minimum uses per activity needed to fill slots
        min_uses = (slots_per_group + num_activities - 1) // num_activities  # ceiling division
        
        if min_uses > self.max_activity_uses:
            old_limit = self.max_activity_uses
            self.max_activity_uses = min_uses
            print(f"Auto-adjusted activity limit from {old_limit} to {self.max_activity_uses}")
            
            # rebuild matrix with new limit and retry solve
            self.matrix = self.generate_weekly_matrix(num_groups)
            self.start_time = time.time()
            if self.solve():
                return True
            else:
                # revert if it still fails
                self.max_activity_uses = old_limit
                return False
        
        return False
    
    def show_error(self, text, title):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.exec_()
    
    def show_info(self, text, title):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.exec_()
    
    def update_group_display(self):
        """Update the group count and members display."""
        num_groups = len(self.groups)
        self.group_count_label.setText(str(num_groups))
        
        if num_groups == 0:
            self.group_members_label.setText("(no groups yet)")
        else:
            member_list = "; ".join([f"{name}: {', '.join(members)}" for name, members in self.groups.items()])
            self.group_members_label.setText(member_list)

    def validate_constraints(self, num_groups: int):
        """Return a list of human-readable constraint issues (empty if OK)."""
        issues = []
        num_activities = len(self.activities)

        if num_activities == 0:
            issues.append("No activities defined")

        if num_groups == 0:
            issues.append("No groups defined")

        # at any given period/day across groups, activities must be unique => need at least num_groups unique activities
        if num_activities < num_groups:
            issues.append(
                f"Too few distinct activities for simultaneous slots: {num_activities} activities < {num_groups} groups (requires at least one distinct activity per group per timeslot)"
            )

        # warn if any group has no participants
        empty_groups = [g for g, members in self.groups.items() if not members]
        if empty_groups:
            issues.append(f"These groups have no participants: {', '.join(empty_groups)}")

        return issues


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())