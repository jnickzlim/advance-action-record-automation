# advance-action-record-automation v1.0.10
Record user action such as mouse clicks and keyboard inputs. Replay repetitively, frequency, cron.

*Remark: Save file in json format.

Features:
- Record
- - Import
  - Save
  - Record
  - Manage
  - Play record, check repeat to repeat multiple times (0=infinity), 
  - Simple coordinate checker
  - Add to replay
  - Add to cron (set time format e.g.[12.15 PM] [05:30 PM])
- Replay
- - Import
  - Save
  - Combine New Import (Lazy to implement - Intention were to combine multiple replay into one - not important)
  - Clear
  - Manage (Edit, Delete, Duplicate, Order arrangement)
  - - Able to set repeat interval, Edit existing actions, set as active or disable
- Cron Job
- - Import
  - Save
  - Clear
  - Manage (Edit, Delete, Duplicate)
  - - Able to edit time, set ative/disable.
    - Select a cron record to play directly 
Shortcuts:
- End key = Pause and resume
- Home key = Start (Multiple press will execute multiple time of the reply)
- Esc key = Stop

##Python: v3.12.2

1. Install the necessary dependancies. (Read the imports at main.py)
2. Run main.py

#Use at your own risk, LOL GLHF
