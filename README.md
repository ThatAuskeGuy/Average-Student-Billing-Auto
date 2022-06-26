# Average-Student-Billing-Auto
Average Student Billing Automation

### What is this?
My current employer is a small SaaS provider in the aviation sector. From the beginning, our billing process has been very hands on. Our developers created a basic billing system page for each customer, but the actual process of invoicing has been very hands on. When I joined the company, we had close to 60 customers, and the average total time it would take to complete a billing cycle was about two weeks depending on the month. Since assisting in billing each month is part of my job description, I know how difficult it is to keep track of every different way we bill our customers based on what they asked for in their contracts.

After six months of doing the billing by hand, I realized there had to be a better way of doing things. I had some previous experience with Python (I took but did not complete a Microsoft course about ML and AI), and so I asked my boss if I could take time away from my other responsibilities in order to try to automate some of the work. He told me I could, and so I spent the next five months diving deep into Python to develop this program. What would normally take us a good two days of formatting Excel spreadsheets to invoice our average student customers (one day if I really tried), now is completed in less than two seconds.

### How does it work?
About a week before billing starts, I get an Excel spreadsheet with all of our customers listed and the invoice number for that month. I take this and input the invoice numbers on an Excel spreadsheet for our customers that we charge based on the average number of students they had the last month. When we start billing, I then feed this file into the program through a GUI i created for it using PySimpleGUI, as well as include a folder that has the raw CSV files for each of our average student billing customers. After less than two seconds, all of our customers have a nicely formatted Excel spreadsheet that we will send out to them.

### What's next?
This is a fairly static program. There is hardly anything about it that changes at all. If the CSV file is ever changed, then I will make changes as necessary to keep it updated, but that rarely happens. Eventually, I will be given the go ahead by our developers to go in to the source code for our billing system and write the logic of this program directly into the billing system's code, completely doing away with a seperate program.
