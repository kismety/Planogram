# Planno — Final Build

## Main workflow
1. Find/select the store.
2. Store Name, Full Address and Store ID are retained.
3. Upload the original text-based PDF planogram OR upload Excel.
4. PDF is converted into the standardized Excel/database structure.
5. Planno uses that structure for:
   - Database
   - Reset checklist
   - Visual planogram
   - Progress
   - PNG
   - Final PDF report

## Gift-card display
Each planogram slot displays:
- Card Name
- Card Value directly underneath

Database fields keep Card Name and Card Value separate for searching/filtering.

## Final one-page PDF
The final PDF uses the same visual layout as the PNG and contains:
- Store Name
- Store ID
- Full Address
- Planogram Name
- Date
- Total items
- Complete
- Missing
- Not Complete
- Completion percentage
- Full C1...Cn / R1...Rn visual planogram
- Card Name + Card Value in each slot

Status:
- Green = Complete
- Yellow = Missing
- Red = Not Complete

The PDF is intentionally scaled to ONE PAGE.
