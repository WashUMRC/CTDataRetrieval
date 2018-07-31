# CTDataRetrieval

This set of scripts is set up to grab all data for a given type of analysis from the Scanco system. It prompts for which machine you're using, the sample number, where to put results, and if you want to gather all of your results (the answer is usually yes). 

Cortical results compiler and trabecularResultsCompilerBulletProof are the scripts that collect standard phenotyping data. Both rely on Excel being installed on the computer being used to gather data, as it uses Excel to parse the text files. These scripts look for key words in the name of the text files generated on the Scanco machines, copy those files over, and parse them, building an output file with only relevant info. Other scripts are put together for specific analyses and should be used by those who know what they're for only.
