# PowerPointFields Add-in
This add-in allows you to reference built-in PowerPoint presentation
properties and any custom properties you define (using VBA) in the
text of your presentation.

## Installation
To install:
1. Open the PowerPointFields.pptm file in PowerPoint.
2. Go to `File/Save As` and save it as a `.ppam` file
3. Go to the `File/Options/Add-ins` dialog.
4. In the `Manage` drop-down, choose `PowerPoint Add-ins` and click `Go...`.
5. Click Add New and you should see `PowerPointFields.ppam`. Open it.
6. Make sure it is active (checked) and it should be active.

## Usage
In any of your shapes that contain text, use `{{field name}}` syntax to
refer to one of the fields from the list below:

* Title
* Subject
* Author
* Keywords
* Comments
* Template
* Last author
* Revision number
* Application name
* Last print date
* Creation date
* Last save time
* Total editing time
* Number of pages
* Number of words
* Number of characters
* Security
* Category
* Format
* Manager
* Company
* Number of bytes
* Number of lines
* Number of paragraphs
* Number of slides
* Number of notes
* Number of hidden Slides
* Number of multimedia clips
* Hyperlink base
* Number of characters (with spaces)
* Content type
* Content status
* Language
* Document version

For example, to include the author's name on a slide, create a text
box with a value of `Author: {{Author}}`.

When you start the presentation, the add-in replaces all field
references with the current document property values. When the
presentation ends, it restores the original text.

## License
The MIT License (MIT)
Copyright © 2021 Scott Means

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
