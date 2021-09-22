# WSUS Reports
![servers](/assets/servers.png)

![pie](/assets/pie.png)

This script was written in order to customize a useful report for compliance and understanding Windows updates with current status.

The script relies on two awesome PowerShell modules that must be installed on the WSUS server. Modules can be installed as follows:

```powershell
Install-Module PoshWSUS
Install-Module ImportExcel
```

Other than that, download this script to your WSUS server and update the following variables, others are optional.

| Variable    | Description                                |
| ----------- | ------------------------------------------ |
| $smtpServer | IP Address of your SMTP Server             |
| $smtpFrom   | From email address                         |
| $smtpto     | Email recipient(s)                         |
| $xlsxPath   | Path location, be sure it's a valid folder |

Now run the script as an Administrator (required for PoshWSUS module).

The script will read data from WSUS, format, create a sheet named 'Servers' listing all servers managed by WSUS. In addition, there will be 4 additional tabs demonstrating a pivot table with various formatting options.



