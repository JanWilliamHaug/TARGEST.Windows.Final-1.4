
<p align="center" width="100%">
    <img width="33%" src="https://user-images.githubusercontent.com/71844869/229735047-6be366d7-8dc8-41f2-bb19-d101691064c0.png">
</p>



# Requirements Tracing Tool

This Python-based requirements tracing tool aims to streamline the process of parsing Word documents, reconciling tags and requirements, and generating various reports for validating the accuracy of the documents and sufficiency of test coverage.

## Features

* Document Parsing: Reads Word documents using the Python/Docx library and extracts paragraphs with embedded tags.
* Tag Extraction: Processes tags at the beginning and end of paragraphs, recognizing parent requirement tags, child requirement tags, and test coverage tags.
* Requirements Hierarchy Construction: Builds a hierarchical structure tree representing the relationships between parent and child requirements.
* Requirement Traceability Analysis: Analyzes the requirements hierarchy and generates various reports, including orphan tags, untested tags, duplicate tags, and validation of parent and child tagging.
* Report Generation: Generates Excel reports using the Python/xlwings library, providing insights into the relationships and traceability of the requirements.
* Graphical Representation: Visualizes the flow of requirements between different documents at a high level using graph visualization libraries like NetworkX or Graphviz.

## Installation

1. Clone the repository:

   git clone [https://github.com/JanWilliamHaug/TARGEST.Windows.Final-1.4.git](https://github.com/JanWilliamHaug/TARGEST.Windows.Final-1.4.git)

2. Install the required libraries:
<br>
   pip install -r requirements.txt

## Usage

Windows: python main.py
macOS: python3 main.py

## Chart

![TARGEST chart](https://user-images.githubusercontent.com/71844869/229740865-bea0329e-c5b3-49a5-acb2-06fe700bf953.png)




