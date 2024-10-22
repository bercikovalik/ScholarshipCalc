# Corvinus University scholarship calculation tool
 This project was developed in response to recent changes in the Regulation on Student Fees and Benefits at my university, requiring a new approach to scholarship calculations. The task was to create a tool that automates the calculation of scholarship scores for students, addressing both complex grouping and redistribution logic while maintaining transparency and accuracy in the process.

# Project Overview
The tool processes student data to calculate a scholarship index (Ösztöndíjindex) based on their academic performance and credits. Students are grouped according to their program name ('KépzésNév'), level of study ('Képzési szint'), and language ('Nyelv ID'—either 'angol' or 'magyar'). They are further classified into year levels based on their active semesters:

1-2 semesters: 1st year
3-4 semesters: 2nd year
5-6 semesters: 3rd year
7-8 semesters: 4th year
Year groups with fewer than 10 students are redistributed to the closest adjacent group. This redistribution process ensures that students from smaller groups are merged with larger groups, running upward or downward depending on availability, to maintain a minimum threshold of 10 students per group. If an entire program has fewer than 10 students, they are excluded from redistribution and processed separately.

# Key Features
Dynamic Grouping and Redistribution: Automatically identifies small groups and redistributes students into adjacent groups while ensuring no loss of student data.
Calculation of Scholarship Scores (KÖDI): For each group, students receive a score between 0 and 100 based on their position relative to the minimum and maximum scholarship index values in their group. The formula used is:
KÖDI = ((HallgÖDI - MinÖDI) / (MaxÖDI - MinÖDI)) x 100

Where:

HallgÖDI is the student's Ösztöndíjindex.
MinÖDI is the lowest Ösztöndíjindex in the group.
MaxÖDI is the highest Ösztöndíjindex in the group.
Output Structure: Generates two output files:
A primary file with updated student year levels and scholarship indices, excluding students from programs with fewer than 10 students.
A separate file containing students from small programs, preserving their original details along with their recalculated year level and Ösztöndíjindex.
This solution not only simplifies the scholarship calculation process but also ensures compliance with the new regulation, creating a fairer and more accurate assessment for all students.

# Technologies Used
Python: For data processing, logic implementation, and automation.
Pandas: To manage and manipulate student data efficiently.
Excel: As the primary format for input and output files, ensuring easy integration with existing university workflows.
Feel free to explore the code and adapt it to similar scholarship distribution scenarios. Contributions are welcome!
