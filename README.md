# Controlling_macro_1
one client's whole Excel VBA codes, uploading data from a Master file into a template.


okay, so you need to know, that this code was created in hungarian language, so variables, names represent names in hungarian. Though, you may find almost every variables defined in the file a0_declare_variables.
Variables are defined in a0, but actually they refer to a master_file, in which the data for the variables are adjusted every month.

Master file:

The master file contains a raw table with the manual and code steps sequenced and with description.
The master file contains the dataset, to which the variables refer.
The master file contains the basic dataset, which are handled and uploaded every month with Excel Power Query.

Basic steps are:
1. Every month receive datasets from the client, and the accountant.
2. adjusting and collecting data needed, eg. categories for costs, if they're not provided correctly.
3. When all data is set, we run the macros in sequence (after other manual data adjustments, like project-cost adjustments).
4. Template is filled with proper data, data gets checked, and explained.
