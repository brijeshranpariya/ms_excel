# Practical Exam – Excel Project

## Overview

This project demonstrates data cleaning, analysis, pivot table reporting, and dashboard creation using Microsoft Excel.

---

# Sheets Included

## Raw_Data

Contains the original dataset used for the project.

## Cleaned_Data

Contains the cleaned and processed dataset.

### Tasks Performed

- Removed inconsistencies
- Organized data properly
- Added conditional formulas

### Formulas Used

```excel
=IF(J2>500,"High","Low")
```

Used to classify sales values.

```excel
=IF(M2>0,"Profit","Loss")
```

Used to identify profit or loss.

---

## Pivot_Tables

Contains pivot tables for analyzing:

- Sales performance
- Profit analysis
- Category-wise insights
- Regional trends

---

## Advanced_Formulas

Contains advanced lookup and analysis formulas.

### Formula Used

```excel
=INDEX(Cleaned_Data!A:A,MATCH(MAX(Cleaned_Data!J:J),Cleaned_Data!J:J,0))
```

Used to find the record with the highest sales value.

---

## Dashboard

Contains a visual dashboard created using charts and summarized data for quick business insights.

---

# Tools Used

- Microsoft Excel
- Pivot Tables
- Charts
- Conditional Formulas
- INDEX & MATCH Functions

---

# Conclusion

This project helped improve practical Excel skills including data cleaning, analysis, formula implementation, and dashboard creation.
