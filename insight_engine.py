def generate_insight(
    analysis_type,
    group_column=None,
    value_column=None,
    operation=None,
    result=None,
    pivot_columns=None
):

    explanation = ""
    insight = ""
    recommendation = ""
    learn = ""
    ai_summary = ""

    # ===========================
    # GROUP BY
    # ===========================

    if analysis_type == "groupby":

        highest_group = result.idxmax()
        highest_value = result.max()

        lowest_group = result.idxmin()
        lowest_value = result.min()

        explanation = f"""
<b>What is Group By?</b><br><br>

Group By organizes your dataset into categories based on
<b>{group_column}</b>.

Instead of looking through every individual record, it combines similar records into groups.

The <b>{operation.upper()}</b> of
<b>{value_column}</b>
was then calculated for each group separately.
"""

        insight = f"""
<b>Key Findings</b><br><br>

• Highest: <b>{highest_group}</b> = <b>{highest_value}</b><br>

• Lowest: <b>{lowest_group}</b> = <b>{lowest_value}</b><br><br>

This allows you to quickly compare how each group performs.
"""

        recommendation = f"""
<b>Recommended Actions</b><br><br>

Study why <b>{highest_group}</b> performs better.

Review the performance of
<b>{lowest_group}</b>
to identify possible improvements.

If the difference is very small, continue monitoring rather than making major changes.
"""

        learn = f"""
<b>When should you use Group By?</b><br><br>

Group By is useful whenever you want to compare categories.

Examples:

🏫 Average score by Class

🏪 Total Sales by Product

⛪ Attendance by Month

🌾 Harvest by Farm

📦 Stock by Category
"""

        ai_summary = f"""
<b>AI Executive Summary</b><br><br>

The dataset was grouped using
<b>{group_column}</b>.

<b>{highest_group}</b>
currently records the highest
<b>{operation}</b>
of
<b>{value_column}</b>.

<b>{lowest_group}</b>
has the lowest value.

This report helps identify the strongest and weakest groups for better decision-making.
"""
    # ===========================
    # PIVOT TABLE
    # ===========================

    elif analysis_type == "pivot_table":

        explanation = f"""
<b>What is a Pivot Table?</b><br><br>

A Pivot Table summarizes large datasets into a simple report.

It groups data using
<b>{group_column}</b>
and compares it across
<b>{pivot_columns}</b>.

The values are calculated using the
<b>{operation.upper()}</b>
function on
<b>{value_column}</b>.
"""

        insight = f"""
<b>Key Findings</b><br><br>

The Pivot Table allows you to compare
<b>{value_column}</b>
across different
<b>{group_column}</b>
and
<b>{pivot_columns}</b>.

This makes trends, high-performing categories, and weak-performing categories much easier to identify.
"""

        recommendation = """
<b>Recommended Actions</b><br><br>

Look for categories with the highest and lowest values.

Investigate why certain groups perform better.

Use these insights to improve planning, budgeting, and decision-making.
"""

        learn = """
<b>When should you use Pivot Tables?</b><br><br>

Pivot Tables are one of the most powerful reporting tools.

Examples:

🏫 Student Scores by Class and Gender

🏪 Sales by Product and Region

⛪ Attendance by Month and Department

🌾 Harvest by Crop and Season

📦 Inventory by Category and Warehouse
"""

        ai_summary = """
<b>AI Executive Summary</b><br><br>

Your Pivot Table successfully summarized the dataset into an easy-to-read report.

Instead of reviewing thousands of individual records, you can quickly compare categories and identify patterns for informed decision-making.
"""  

    elif analysis_type == "statistics":

        if operation == "mean":

            explanation = """
The Mean is the average value of all numbers in your dataset.

It is calculated by adding all values together and dividing by the total number of values.
"""

            insight = """
The Mean helps you understand the typical value in your dataset.

If the Mean is much higher or lower than expected, there may be unusually large or small values affecting the result.
"""

            recommendation = """
Compare the Mean with the Median.

If they are very different, your data may contain outliers.
"""

            learn = """
The Mean is one of the most common statistical measures.

Businesses use it to calculate average sales.

Schools use it to calculate average scores.
"""

            ai_summary = """
Your dataset has been summarized using the Mean.

This gives a quick overview of the average performance.
"""

        elif operation == "median":

            explanation = """
The Median is the middle value after arranging all numbers from smallest to largest.
"""

            insight = """
The Median is not affected by extreme values.

It often represents the true center of your data better than the Mean.
"""

            recommendation = """
If your dataset contains very large or very small values, use the Median instead of the Mean.
"""

            learn = """
Median is useful for income, salaries, house prices and exam scores where outliers exist.
"""

            ai_summary = """
The Median shows the middle point of your dataset.
"""

        elif operation == "mode":

            explanation = """
The Mode is the value that appears most frequently.
"""

            insight = """
The Mode helps identify the most common value or category.
"""

            recommendation = """
Use the Mode to identify popular products, common grades or frequent attendance.
"""

            learn = """
Mode is especially useful for categorical data.
"""

            ai_summary = """
The most frequent values have been identified.
"""

        elif operation == "sum":

            explanation = """
The Sum adds together all numeric values.
"""

            insight = """
This shows the total amount represented in your dataset.
"""

            recommendation = """
Use totals to measure revenue, expenses, attendance or inventory.
"""

            learn = """
Businesses commonly use Sum to calculate total sales and profits.
"""

            ai_summary = """
The total values have been calculated successfully.
"""

        elif operation == "count":

            explanation = """
Count shows how many records exist.
"""

            insight = """
This tells you the size of your dataset.
"""

            recommendation = """
Always verify the number of records before beginning analysis.
"""

            learn = """
Count measures the total number of observations.
"""

            ai_summary = """
The total number of records has been calculated.
"""

        elif operation == "min":

            explanation = """
Minimum is the smallest value in the dataset.
"""

            insight = """
This identifies the lowest-performing observation.
"""

            recommendation = """
Investigate why the minimum value is significantly lower than others.
"""

            learn = """
Minimum helps detect poor performance.
"""

            ai_summary = """
The smallest value has been identified.
"""

        elif operation == "max":

            explanation = """
Maximum is the largest value in the dataset.
"""

            insight = """
This identifies the highest-performing observation.
"""

            recommendation = """
Study why the highest value performs better than others.
"""

            learn = """
Maximum helps identify top performers.
"""

            ai_summary = """
The highest value has been identified.
"""

        elif operation == "std":

            explanation = """
Standard Deviation measures how spread out the data is.
"""

            insight = """
A small Standard Deviation means values are close together.

A large one means values vary greatly.
"""

            recommendation = """
Investigate datasets with high variation.
"""

            learn = """
Standard Deviation measures consistency.
"""

            ai_summary = """
The variation of your dataset has been calculated.
"""

        elif operation == "var":

            explanation = """
Variance measures how much values differ from the average.
"""

            insight = """
Higher variance means greater variability.
"""

            recommendation = """
Compare Variance across different groups.
"""

            learn = """
Variance is closely related to Standard Deviation.
"""

            ai_summary = """
The spread of your data has been measured.
"""

        else:

            explanation = f"""
The {operation.title()} statistic has been calculated successfully.
"""

            insight = """
The selected statistic provides additional information about the distribution of your data.
"""

            recommendation = """
Combine this statistic with charts for deeper insights.
"""

            learn = """
Statistical analysis helps understand patterns hidden inside data.
"""

            ai_summary = f"""
Your dataset has been analyzed using {operation.title()}.
"""



    return {
        "explanation": explanation,
        "insight": insight,
        "recommendation": recommendation,
        "learn": learn,
        "ai_summary": ai_summary
    }