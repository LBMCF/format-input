format-input
======================
[![License](https://poser.pugx.org/badges/poser/license.svg)](./LICENSE)

This script reads the exported (.csv|.txt) files from [Scopus](https://www.scopus.com), [Web of Science](https://clarivate.com/webofsciencegroup/solutions/web-of-science), [PubMed](https://www.ncbi.nlm.nih.gov/pubmed), [PubMed Central](https://www.ncbi.nlm.nih.gov/pmc) or [Dimensions](https://app.dimensions.ai) databases and turns each of them into a new file with an unique format. This script will ignore duplicated records.

## Table of content

- [Pre-requisites](#pre-requisites)
    - [Python libraries](#python-libraries)
- [Installation](#installation)
    - [Clone](#clone)
    - [Download](#download)
- [How To Use](#how-to-use)
- [Author](#author)
- [Organization](#organization)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## Pre-requisites

### Python libraries

```sh
  $ sudo apt install -y python3-pip
  $ sudo pip3 install --upgrade pip
```

```sh
  $ sudo pip3 install argparse
  $ sudo pip3 install xlsxwriter
  $ sudo pip3 install pandas
  $ sudo pip3 install colorama
```

## Installation

### Clone

To clone and run this application, you'll need [Git](https://git-scm.com) installed on your computer. From your command line:

```bash
  # Clone this repository
  $ git clone https://github.com/LBMCF/format-input.git

  # Go into the repository
  $ cd format-input

  # Run the app
  $ python3 format_input.py --help
```

### Download

You can [download](https://github.com/LBMCF/format-input/archive/master.zip) the latest installable version of _format-input_.

## How To Use

```sh
  $ python3 format_input.py --help
  usage: format_input.py [-h] -t {scopus,wos,pubmed,pmc,dimensions,txt} -i
                         INPUT_FILE [-o OUTPUT] [--version]

  This script reads the exported (.csv|.txt) files from Scopus, Web of Science,
  PubMed, PubMed Central or Dimensions databases and turns each of them into a
  new file with an unique format. This script will ignore duplicated records.

  optional arguments:
    -h, --help            show this help message and exit
    -t {scopus,wos,pubmed,pmc,dimensions,txt}, --type_file {scopus,wos,pubmed,pmc,dimensions,txt}
                          scopus: Indicates that the file (.csv) was exported
                          from Scopus | wos: Indicates that the file (.csv) was
                          exported from Web of Science | pubmed: Indicates that
                          the file (.csv) was exported from PubMed | pmc:
                          Indicates that the file (.txt) was exported from
                          PubMed Central, necessarily in MEDLINE format |
                          dimensions: Indicates that the file (.csv) was
                          exported from Dimensions | txt: Indicates that it is a
                          text file (.txt)
    -i INPUT_FILE, --input_file INPUT_FILE
                          Input file .csv or .txt
    -o OUTPUT, --output OUTPUT
                          Output folder
    --version             show program's version number and exit

  Thank you!
```

## Author

* [Glen Jasper](https://github.com/glenjasper)

## Organization
* [Molecular and Computational Biology of Fungi Laboratory](https://sites.icb.ufmg.br/lbmcf/index.html) (LBMCF, ICB - **UFMG**, Belo Horizonte, Brazil).

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.
