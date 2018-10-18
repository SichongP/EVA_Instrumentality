mkmetadata can read information from three optional config files when autofilling metadata file. Currently only mandatory fields in metadata (see FAQ) is supported.
The basic format for each config file is:
```
ATTRIBUTE="value"
```
For example, a sample user_info.config file may look like this:
```
LAST_NAME="Franklin"
FIRST_NAME="Rosalind"
TELEPHONE="1234567890"
EMAIL="refranklin@doublehelix.com"
LABORATORY="Franklin Lab"
CENTER="King's College London"
ADDRESS="London WC2R 2LS, UK"
```

mkmetadata looks for ```analysis_info.config``` and ```project_info.config``` under each analysis and project folder when scanning files.
Thus, if you wish to use config files to fill out metadata rather than mannually input them afterwards, you can simply make a config file and put it under corresponding project/analysis folder.

A sample project_info.config may look like this:
```
TITLE="Unveiling DNA structure"
ALIAS="dsDNA"
DESCRIPTION="Unprecedented discovery of code of life"
CENTER="King's College London"
TAXID="9606"
```

A sample analysis_info.config may look like this:
```
TITLE="X Ray Crystallography of Deoxyribonucleic acid"
DESCRIPTION="Using X Ray Crystallography reveals structure of Deoxyribonucleic acid"
ALIAS="CrystalDNA"
EXPERIMENT_TYPE=""
REFERENCE=""
REFMD5=""
```
Below is a list of analysis attributes currently allowed in metadata file:
```
(Required)
TITLE=""
ALIAS=""
DESCRIPTION=""
EXPERIMENT_TYPE=""
REFERENCE=""
REFMD5=""
(Optional)
PLATFORM=""
SOFTWARE=""
PIPELINE=""
IMPUTATION=""
PHASING=""
CENTER=""
DATE=""
ANAL_LINK=""
RUN=""
```

And here is a list of project attributes currently allowed in metadata file:
```
(required)
TITLE=""
ALIAS=""
DESCRIPTION=""
CENTER=""
TAXID=""
(Optional)
PUBLICATION=""
PARENT=""
CHILD=""
PEER=""
PRJ_LINK=""
HOLD_DATE=""
COLLABORATORS=""
STRAIN=""
BREED=""
BROKER=""
```
