# probabilistic_color_code
Probabilistically color single letters in Word documents

Run from main directory containing "color_text_consistency.py".
The script will prompt asking whether you want the colors and the font changed.

The directory /colors/ should contain:
rgb_colors.csv
probability_distributions_v1.csv
probability_distributions_v2.csv
sub-xxx_prediction_space.csv (for all subjects)

There are two groups of letters, counterbalanced across participants (trainees).
Letter group 1: e	s	m	q	x	c	h	o
Letter group 2: 

In the probability distribution CSV files, the letter identity row=column, should correspond to that letter's consistency category. 
For instance, if "x" is assigned to the 100% consistency category, then for row=x=column, should have the value of "1" and all other columns should have a value of "0".
If "s" is assigned to the 75% consistency category, then row=x=column, should have the value of "0.75" and all other columsn should have a value of "0.25/(number of letters-1)".
All columns should sum to 1 (probability distribution)!

The file "sub-xxx_prediction_space.csv"  will indicate while letter is assigned to which consistency category for that participant, and also which colorcode. 