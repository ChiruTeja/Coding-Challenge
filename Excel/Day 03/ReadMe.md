Excel Day 03 -- Hospital Patient Service Analysis

Name: ArunaTeja \
GitHub username: https://github.com/ChiruTeja/    \
\
Method\
DataSet Understanding \
The source workbook contains Appointment_Data worksheet with 24 Appointment Details and 7 fields. Each Row Represents one unique Appointment.

<img width="806" height="161" alt="image" src="https://github.com/user-attachments/assets/e8c6dc86-8aa7-4fae-ac86-bea70fcb9f4e" />


Workbook preparation \
The dataset contains 24 appointment records. The original Appointment_Data worksheet is retained unchanged. A separate Service_Analysis worksheet is created, and the data is converted into an Excel Table named HospitalAppointments.

BUSINESS RULES \
•15 minutes or less: Within Target.
•16–30 minutes: Minor Delay.
•31–45 minutes: Major Delay.
•More than 45 minutes: Critical Delay.
•Escalate if wait time is greater than 45 minutes OR feedback is Complaint.
•Priority Follow-up if wait time is greater than 30 minutes AND feedback is Complaint.

CALCULATED COLUMNS \
The following calculated columns were added to Service_Analysis:

1. Wait_Category – IFS

Waiting-time categories are calculated using the business rules provided.\
Excel formula: \
Wait_Category=IFS([Wait_Time_Minutes]<=15,"With in Target",[@[Wait_Time_Minutes]]<=30,"Minor Delay",[@[Wait_Time_Minutes]]<=45,"Major Delay",[@[Wait_Time_Minutes]]>45,"Critical Delay")

2. Escalation_Status – IF with OR

An appointment should be escalated when:

Waiting time is greater than 45 minutes, OR

Feedback is Complaint.

=IF(OR(E2>45,G2="Complaint"),"Escalate","No Escalation")

OR returns TRUE if at least one condition is satisfied.

3. Follow_Up_Priority – IF with AND

Priority follow-up is required when:

Waiting time is greater than 30 minutes, AND

Feedback is Complaint.

=IF(AND(E2>30,G2="Complaint"),"Priority Follow-up","Standard Follow-up")

AND returns TRUE only when both conditions are satisfied.
