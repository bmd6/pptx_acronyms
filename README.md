# pptx_acronyms

Basic Usage (without known acronyms or exclusions):<br>
`python acronym_finder.py path_to_presentation.pptx`

With Known Acronyms:<br>
`python acronym_finder.py path_to_presentation.pptx --known-acronyms known_acronyms.csv`

With Known Acronyms and Exclusions:<br>
`python acronym_finder.py path_to_presentation.pptx --known-acronyms known_acronyms.csv --exclude-acronyms exclude_acronyms.csv`

With Custom Log Level:<br>
`python acronym_finder.py path_to_presentation.pptx --log-level DEBUG`
