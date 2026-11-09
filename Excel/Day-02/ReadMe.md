Excel Day 02 -- Loan Portfolio Analysis
Student details
- Name: ArunaTeja
- GitHub username: https://github.com/ChiruTeja/ 
 
Method
1. DataSet Understanding \
   The source workbook contains Loan_Data worksheet with 20 Loan Account Details and 6 fields. Each Row Represents one unique Loan account.
   
   <img width="568" height="141" alt="image" src="https://github.com/user-attachments/assets/f740e41b-972d-423f-8219-c63c4551fca7" />
2. Workbook preparation \
   The Original Worksheet Loan_Data unchanged, created a seperate Loan_Analysis worksheet and converted the copied source data into Excel      table named 'LoanPortfolio'
3. Calculated columns\
The following calculated columns were added to Loan_Analysis:
    - Outstanding_Balance: Original loan amount less amount repaid.\
         Outstanding_Balance=Loan_Amount_INR-Amount_Repaid_INR
    - Repayment_Percentage: Amount repaid divided by original loan amount. The result is formatted as a percentage.\
         Repayment_Percentage=Amount_Repaid_INR/Loan_Amount_INR
    - Portfolio_Share: Individual original loan amount divided by the total original loan amount.\
         Total_Amount=SUM(Loan_Amount_INR)\
         Portfolio_Share=Loan_Amount_INR/Total_Amount
 4. Assumption and illustrative interest \
     In the Assumptions worksheet, I entered a hypothetical annual simple-interest benchmark of 10.5%. The illustrative annual interest
     amount is calculated on each original principal using an absolute reference to the assumption cell (shown here as B2):\
      Annual_Interest_Estimate=Loan_Amount_INR*Assumptions!$B$2\
        This is an illustrative annual simple-interest estimate, not actual  interest.
 5. Summary calculations\
     The summary uses Excel functions including SUM, AVERAGE, MIN,MAX, COUNT/`COUNTA`, and ROUND. Example formulas include:\
         =SUM(LoanPortfolio[Loan_Amount_INR])\
         =SUM(LoanPortfolio[Amount_Repaid_INR])\
         =SUM(LoanPortfolio[Outstanding_Balance])\
         =AVERAGE(LoanPortfolio[Loan_Amount_INR])\
         =MIN(LoanPortfolio[Loan_Amount_INR])\
         =MAX(LoanPortfolio[Loan_Amount_INR])\
         =COUNTA(LoanPortfolio[Loan_ID])\
         =ROUND(SUM(LoanPortfolio[Amount_Repaid_INR])/SUM(LoanPortfolio[Loan_Amount_INR]),4)
6. Cell references
   - Relative reference: E2 changes when the formula is copied to another row.
   - Absolute reference: $B$2 keeps both the column and row fixed when copied.
   - Dollar signs are used so the assumption or total reference does not shift when a formula is filled down.
   - Mixed references: $A1 locks the column but allows the row to change; A$1 locks the row but allows the column to change.
7. Key findings
    1. The dataset contains 20 loan accounts. Total original principal disbursed is ₹3,50,25,000, total repayments are ₹1,52,80,000, and total            outstanding balance is ₹1,97,45,000. Overall repayments represent approximately 43.63% of the original principal.
    2. The East Region has the highest outstanding balance of 70.95 Lakhs, representing approximately 35.93% of the total outstanding.
    3. Home loans contribute the highest outstanding balance of 11.46 Lakhs,  approximately 58.04% of the total outstanding loan portfolio.
      The largest individual outstanding account is LN3007, an East-region Home loan, with ₹39,00,000 outstanding.
8. Summary metrics

    <img width="340" height="161" alt="image" src="https://github.com/user-attachments/assets/7823d8e4-394d-4c38-9f0c-8e2adf07f12c" />   
10. Recommendation\
     The bank should closely monitor high value outstanding home loans, perticularly accounts with low repayment percentages. 
     It should prioritise repayment follow-ups in East region and perform regular reviews on the portfolio to manage credit risks.
11. Limitations / assumptions\
     The dataset contains only 20 loan ID's and 6 columns that should not include Loan_tenure, Interest_Rate, EMI schedules,  Customer_Credit_scores,    Overdue_Days.
   So, it is not sufficient to actual Loan performance analysis.
   
Files
- Day02_Loan_Analysis.xlsx

   

 






 
