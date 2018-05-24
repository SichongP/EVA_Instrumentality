1. What does mkmetadata do?  
mkmetadata reads files you provide and fill in some information required in metadata file to submit to EVA. mkmetadata use your file structure to determine your project structure. See [README](https://github.com/SichongP/EVA_Instrumentality/blob/master/README.md) for more details.
2. What do I need for variant submission to EVA?  
You will need your variant files (vcf, bed, wig, etc.) and a metadata file with same format as [EVA template](https://github.com/SichongP/EVA_Instrumentality/blob/master/src/EVA_Submission_template.V1.1.0.xlsx) 
3. What do I need to run mkmetadata?  
You will need your variant files, organized as instructed in [README](https://github.com/SichongP/EVA_Instrumentality/blob/master/README.md). Optionally, you can provide config files for your Submitter information, project information, and analysis information so mkmetadata can autofill these fields for you. See [here](https://github.com/SichongP/EVA_Instrumentality/blob/master/config_guidline.md) for details on how to format your own config files.
4. How does mkmetadata recognize my projects?  
mketadata takes a path as input. Any folder under that path is considered a project. Any folder under each project folder is considered an analysis associated with that project. Any files under each analysis folder are associated to said analysis.  
If a project_info.config or analysis_info.config file is found under a project or analysis folder, mkmetadata will try to read information about that project or analysis and autofill correspoding fields in output metadata file. [See here](https://github.com/SichongP/EVA_Instrumentality/blob/master/config_guidline.md) for details on how to format your own config files.
5. Can I upload files that are not associated with any specific analysis?  
No. EVA guidline requires every file in metadata be associated with at least one analysis
6. What is reference in an analysis?  
The reference assembly against which the analysis was performed. You can either provide a url to the reference file hosted on a publically available server (NCBI, ensemble, etc) or an ENA accession. For human referencec, a simple GRC reference name is accepted. 
7. What is MD5sum of reference?  
MD5sum is a hash sequence calculated from content of a file. Although vulnerable to intentional corruptions and attacks, MD5sum is widely used to ensure files are free of unintentional corruptions and used by EVA to ensure correct version of reference is associated with an analysis. If you modified your reference by any means, you must calculate its own MD5sum value. To calculate MD5sum value of a file, run ```md5sum reference.fa``` on unix machine (Linux or Mac OS) or ```CertUtil -hashfile editdistance.sh MD5``` on Windows machine. You also need to provide a publically accessible link to your modified reference. (see FAQ #6)  
If you did not alter a reference, you can usually get this value from public database that houses the reference. [See here for an example.](http://ftp.ncbi.nlm.nih.gov/genomes/all/GCA/000/001/405/GCA_000001405.27_GRCh38.p12/md5checksums.txt)
8. 
