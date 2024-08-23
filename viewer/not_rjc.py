"""
Author: Manuel Rosa
Send windows notifications when a new project is added or alerts in a set time interval
when a project needs to be reviewed
"""

import os
import sys
import pandas as pd
from plyer import notification
import time


class notification_app:
    def __init__(self):
        self.script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        options = os.path.join(self.script_dir, "options.txt")
        self.df_folders = pd.DataFrame(
            columns=["Conjunto", "Utilizador", "Data Edição", "Verif.", "Version"]
        )
        with open(options, "r") as file:
            lines = file.readlines()
        for line in lines:
            if "phc =" in line:
                folder_path_strip = line.split("= ", 1)[1].strip()
            if "time =" in line:
                time = (line.split("= ", 1)[1]).strip()
        self.folder_path = os.path.join(folder_path_strip, "Obras")
        self.time_per_update = 10
        self.time_per_notification = int(time)
        self.time_loops_max = self.time_per_notification / self.time_per_update
        self.start_loop = self.time_loops_max - 1
        self.start = 1
        self.list_notifications()
        self.no_noti = 0

    def walklevel(self, some_dir, level):
        """Gives me subfolders x levels beyond the main folder,
        used to get the actual projects that are subfolders inside of multiple folders in the main path
        """
        some_dir = some_dir.rstrip(os.path.sep)
        assert os.path.isdir(some_dir)
        num_sep = some_dir.count(os.path.sep)
        for root, dirs, files in os.walk(some_dir):
            yield root, dirs, files
            num_sep_this = root.count(os.path.sep)
            if num_sep + level <= num_sep_this:
                del dirs[:]

    def list_notifications(self):
        while True:
            unique_values = pd.DataFrame()
            old_df_folders = self.df_folders
            self.df_folders = pd.DataFrame(
                columns=["Conjunto", "Utilizador", "Data Edição", "Verif.", "Version"]
            )  
            folders = self.walklevel(self.folder_path, level=2)
            list_folder = []
            paths = []
            paths_list = []
            max_path_len = 0
            project_paths = []
            for folder in folders:
                list_folder.append(folder[0])
            for (
                pathies
            ) in (
                list_folder
            ):  
                paths = pathies.split("\\")
                if len(paths) > max_path_len:
                    max_path_len = len(
                        paths
                    )  
                paths_list.append(paths)  
            project_paths = [
                x for x in paths_list if len(x) >= max_path_len
            ]  
            project_paths = [
                i for n, i in enumerate(project_paths) if i not in project_paths[:n]
            ]  
            project_paths_real = (
                []
            )  
            for (
                paths
            ) in (
                project_paths
            ):  
                paths[0] = paths[0] + "\\"
                path = os.path.join(*paths)
                project_paths_real.append(path)
            for (
                path
            ) in (
                project_paths_real
            ):  
                log_path = os.path.join(path, "log.txt")
                if os.path.exists(log_path):
                    with open(log_path, "r") as f:
                        lines = f.readlines()

                name_project = (
                    path.split("\\")[-2] + ": " + path.split("\\")[-1]
                )  
                verified = None
                if lines[-1].startswith("V"):
                    verified = 0
                elif lines[-1].startswith("."):
                    verified = 2
                else:
                    verified = 1

                last_version = lines[-1].split(":")[0]
                last_user = (
                    (lines[-1].split(":"))[1].split(",")[0]
                ).split()  
                last_date = ((lines[-1].split(":")[1]).split(",")[1])[
                    1:11
                ]  
                list_df = [
                    name_project,
                    last_user,
                    last_date,
                    verified,
                    last_version,
                ]  
                if verified == 0:
                    self.df_folders = self.df_folders._append(
                        pd.Series(list_df, index=self.df_folders.columns),
                        ignore_index=True,
                    )  

            if len(self.df_folders) == 1:
                for index, row in self.df_folders.iterrows():
                    desig = row["Conjunto"]
                    author = row["Utilizador"][0]
                    version = row["Version"]
                title_noti = f"{desig} - {version}"
                message_noti = f"Obra pendente por verificar"
                self.no_noti = 1
            elif len(self.df_folders) > 1:
                title_noti = f"Várias obras por rever"
                message_noti = f"Obras pendentes por verificar"
                self.no_noti = 1
            else:
                self.no_noti = 0
            self.start_loop = self.start_loop + 1
            time.sleep(self.time_per_update)
            if self.start == 1:
                old_df_folders = self.df_folders
                self.start = 0

            print(self.df_folders)
            print(old_df_folders)

            unique_df = pd.concat([self.df_folders, old_df_folders])
            columns = ["Conjunto", "Version"]
            unique_values = unique_df.drop_duplicates(subset=columns, keep=False)
            unique_values = unique_values.dropna(how="all")

            if len(self.df_folders) > len(old_df_folders):
                for index, row in unique_values.iterrows():
                    desig = row["Conjunto"]
                    version = row["Version"]
                    notification.notify(
                        title=f"{desig} - {version}",
                        message=f"Obra adicionada",
                        timeout=5,
                    )
            print(f"Notificação:{self.no_noti}")
            print(
                f"startloop = {self.start_loop} and time_loop_max = {self.time_loops_max}"
            )
            if (self.start_loop == self.time_loops_max and self.no_noti == 1) or (
                self.start_loop > self.time_loops_max and self.no_noti == 1
            ):
                self.start_loop = 0
                notification.notify(
                    title=title_noti,
                    message=message_noti,
                    timeout=5,
                )


notification_app()
