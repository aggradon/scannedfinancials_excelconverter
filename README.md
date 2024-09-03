<Broad Definition>
I want to build a tool that leverages GPT vision models to extract financials from scanned documents provided by the user, and output in a templatised model sequential 5-10Ys of adjoined PnL.
</Broad Definition>

<PnL Production Specific Definition>
    - User will input 1-10 years worth of scanned statements in images
    - The tool will call GPT vision api to produce transcriptions of the data in a tab delimited format (or any other format copyable to excel)
    - the tool then iteratively sends the next image to add to the line items of the initialised Income statement. This is done iteratively to avoid GPT missing line items as name changes/additions/subtractions in line items across the years occurs.
    - the tool then checks the transcription at summation lines for accuracy (for eg. Gross Profit is Revenue - COGS. The tool does that summation and checks for any wrongly transcribed years data and fixes it)
    - then exports it into an excel file, writes some CAGR formulas into cells and does some formatting (alighnemnt, borders, and fills)
    - exports the excel workbook
<PnL Production Specific Definition>
    
