# EVA_Instrumentality
This program is aimed at helping researchers submit to European Variant Archive with ease by aiding completion of metadata required by the archive.

## Issues, bugs :bug:, and new feature requests
Please open an issue in this repository. When reporting an issue, please provide sufficient information to reproduce the error. Some examples include:

* Full error message
* Files associated with error message
* OS information (Linux/Mac OS/Windows)

## What you will need:
* All your variant files (vcf, gff, bed, etc.) organized as instructed below
* *python2*
* python module *openpyxl*
  * To install openpyxl, run the following code:
     ```shell
     pip install openpyxl
     ```
  * You can check if you have *openpyxl* available by doing the following:
    ```python
    python2
    >>> from openpyxl import Workbook
    ```
  * Once you have *openpyxl* installed, this should run without errors issued.
  * Visit [openpyxl website](https://openpyxl.readthedocs.io/en/stable/) for more information and help
* (Optional) user_info.config file
* (Optional) project_info.config file
* (Optional) analysis_info.config file


## Get started:

1. Copy this repository to your local machine:
```shell
git clone https://github.com/SichongP/EVA_Instrumentality
cd EVA_Instrumentality
```

2. **Follow the instruction in the next section to organize your files.**

3. Run mkmetadata:
```
chmod +x mkmetadata
./mkmetadata --out output.xlsx path_to_directory_containing_all_projects
```
mkmetadata will scan all folders under ```path_to_directory_containing_all_projects``` you provided and treat each folder as a project. So make sure you have all project folders under one single folder and no other folders in there

4. Additionally, you can provide a configuration file containing your information so mkmetadata can autofill "Submitter Details" sheet. A template config file can be found at src/user_info.config. [See here]("https://github.com/SichongP/EVA_Instrumentality/blob/master/config_guidline.md") for more information.
To include a user information file, use ```--user user_info.config```

5. Complete the rest of metadata file. (See FAQ for help)


## Organize your files:

The program relies on your file structure to determine the relationships between your projects, analyses, and files. Thus, it is crucial that you organize your files correctly for the program to generate correct metadata file.

*Alternatively, you can refer to FAQ to make your own metadata file if you do not wish to re-organize your files*

You should have a folder for each project. 

All analyses associated with such project should be put in their own folder under the project folder. 

Variant files associated with each analysis should be put under the analysis folder.

info.config file should be put directly under the project/analysis folder they belong to.

See below structure (trailing slash ```/``` indicates a directory instead of a file)

* Projects/
  * project1/
    * project_info.config (optional)
    * analysis1/
      * analysis_info.config (optional)
      * file1_1.vcf
      * file1_2.vcf
    * analysis2/
      * analysis_info.config (optional)
      * file2_1.vcf
      * file2_2.vcf
  * project2/
    * project_info.config (optional)
    * analysis3/
      * analysis_info.config (optional)
      * file3_1.vcf
      * file3_2.vcf
    * analysis4/
      * analysis_info.config (optional)
      * file4_1.vcf
      * file4_2.vcf

