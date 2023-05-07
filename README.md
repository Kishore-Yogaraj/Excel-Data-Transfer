# Excel-Data-Transfer
This program moves data from a specified cell in a source excel file to a specified cell in a destination excel file whenever the source cell is changed/updated

## Inspiration For the Application

During my internship with Magna, a leading mobility technology company, I had the privilege of working with the General Motors team and managing a variety of tasks. One of my key responsibilities was managing the company's inventory trackers using Excel files, which needed to be updated on a weekly basis with the quantities of parts received and required for the next shipment. However, I noticed a problem with the EDI list, which is the master Excel sheet provided by General Motors - the quantities for the 82 different parts were constantly changing, making it difficult for the team to accurately track the required parts.

To solve this problem, I took the initiative to create a Python program that enables a user to pair a cell from the source file (EDI list) to a cell in the destination file (inventory tracker). This innovative solution ensures that every time the source cell changes, the destination cell will automatically adjust accordingly. This means that the inventory tracker can now be kept 100% accurate, without the need for manual inputting of required quantities.

With this program in place, the entire engineering team can now easily keep track of which parts need to be run in the cell or put into press. This not only saves time and effort but also greatly improves the efficiency and accuracy of the entire inventory management process. Overall, I am proud to have contributed to Magna's success by creating a tool that streamlines their inventory management system, and enhances the accuracy and efficiency of the entire process.

## Demonstration (Excel_Data_Transfer.webm)

During this demonstration, I showcase the efficiency and accuracy of my program by highlighting the parts and quantities that we will be working with. The file I use represents the "EDI" list or master file, which can also be referred to as the "source file". To conduct the demonstration, I created a copy of the original file used by General Motors for illustrative purposes.

In the next Excel file, I highlight the corresponding quantities and parts that we will be working with. This file is commonly referred to as the "destination file" or "inventory tracker", which is a replica of the file used by the engineering team to determine the quantities needed for a shipment. For the purpose of the demonstration, I only worked with five different parts, but the program is designed to work seamlessly with all 82 parts, as long as the cell from the source file is paired with the cell in the destination file.

To demonstrate the effectiveness of the program, I go back to the source file and intentionally change the quantities, as if General Motors had randomly changed them. Once the changes have been made, I close the destination file and run the program. It takes approximately 30 seconds for the program to run, and once completed, I open the destination file to reveal that the quantities have been automatically updated thanks to the program.

This demonstration proves the reliability of the program and its ability to ensure that the inventory tracker matches the source file accurately and efficiently, without the need for manual intervention.
