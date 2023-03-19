**Meal Plan Generator**

This tool is used to generate a maximally diversified meal plan, based on your own meal collection, while adhering to certain rules. It targets one cooked meal per calendar day.

The implementation via Excel with Visual Basic macros has the following advantages and disadvantages:

Advantages

- no user interface needs to be programmed – the user interface is Excel
- simple data storage, sorting, filtering, etc. with the Excel tools on board
- VBA as a simple programming language.

Disadvantages

- only usable with Excel
- not cloud-enabled
- VBA as a relatively primitive language with clear disadvantages compared to professional languages like C++, C#, etc.

The Excel workbook is divided into four worksheets: Meals, Rules, Planning, and Language.

**Worksheet Meals**

It contains the collection of all meals which we can cook.

Several attributes are defined for every meal. Some of these are used during the automatic planning, some are potentially useful or interesting.

The meal collection is stored in a filterable table. The meals are in the rows and their attributes in the columns. If necessary, selected columns can be easily hidden or shown through the Excel standard functionality in order to improve the readability of the table. When hidden, the data is retained and will still be used.

Meaning of the columns:

- **Meal**. Contains the name of the meal. The name must be unique for every meal.
- **Variants / Side meals**. Additional information on possible variants & side meals. For example, we do not differentiate between "spaghetti with tomato sauce" and "penne with tomato sauce". Both are just variants of the meal "pasta with tomato sauce".
- **Category**. The category of the meal. This is essential for planning, as will become clear later on the worksheet Rules.
- **Last serving date**. The date when this meal was last cooked. Is empty initially, but becomes increasingly important with a longer use. It is used to maximize the repetition distance for each meal. This ensures a maximally varied meal plan.
- **Prior serving dates**. This text field collects all previous serving dates. It is potentially useful for archiving or analysis purposes, e.g. if someone wishes to do a statistical analysis of their own meal history. Otherwise, it is not actively used.
- **Erase dates**. This button is located above the "Last serving date" column. It clears all dates in the two date columns. This is very useful for initial playing-around and experimenting with the tool. Later, during the real use, the function should no longer be used in order to avoid data loss.
- **Effort**. An integer: from 1 for lowest cooking effort to X for highest cooking effort. We use the scale from 1 to 3 (1 - easy, 2 - moderate, 3 - laborious), but if you wish, you can also use higher-level scales, e.g. from 1 to 10. The effort plays an important role during the automatic planning, see the worksheet Rules.
- **Popularity**. An integer that indicates how popular the meal is among the family members, similar to a star rating. We use the following scale, but you can also define others.

0. Excluded. We never cook it. This value can be used to keep some meals in the collection but exclude them from the planning.
1. Controversial. Some family members do not eat this meal at all.
2. OK. At least acceptable for all family members.
3. Undisputed. Okay for all family members, popular with some.
4. Favorite. Popular with most or all family members

The popularity plays an important role during the automatic planning, see the worksheet Rules.

- **Ingredients**. Which ingredients (which we don't always have at home) do you have to buy for this meal? These ingredients are used to generate the shopping lists. We assume that we always have some basic meals in stock at home, such as pasta, tinned tomatoes, salt, etc. Therefore, we do not include such basic ingredients here.
- **Contains meat**. An integer that indicates whether (or how strongly) the meal contains meat. We use the following scale:

0. a vegetarian meal with no meat contained
1. may contain some meat in parts or optionally (e.g. homemade pizza may have salami in some sectors, other sectors may be vegetarian; cooked lentils may be served optionally with or without sausage)
2. the meat is essential for this meal.

This attribute is not used during the automatic planning, but is very useful for an ad hoc inspiration.

- **Other attributes**. Hopefully these are self-explanatory. They are currently not used during the automatic planning, but can be of interest for an ad hoc inspiration.

You can also insert **your own additional columns** and fill them with your own data. (E.g. columns "Recipe URL", "Calories", etc.) Additional columns may **only** be added to the right of the original columns - not to the left or in the middle.

**Use Case "Ad Hoc Inspiration"**

- What could I cook today? It should be something vegetarian, not sweet, and all family members should like it, or at least not dislike it. Simply filter the columns of the table accordingly: "Contains meat" = 0; "Category " ! = Sweet dish; "Popularity" \> 1.
- I have some bell peppers left. What could I use them for? Simply filter in the table.
 Column Filtering: "Ingredients" contains "bell pepper".
- Which fish meal haven't we had for a long time? Simply filter & sort in the table.
 Column Filtering: "Category" = Fish ; Column sorting: "Last served date" - ascending.

**Worksheet Rules**

Here we define the rules for the automatic meal plan generation.

- **Shopping day**. Indicates on which day of the week the big weekly grocery shopping is typically done. Logic: The ingredients for the meal that is cooked on the shopping day are also bought that day. (I.e. the shopping takes place on this day _before_ thecooking.) Hence, if you buy do your groceries e.g. on Friday evening, then you should set your shopping day in the tool to Saturday. Otherwise the purchases for the Friday meal would come too late...
- **Maximum meal effort on a weekday**. Meals that cause more work than this value are considered to be "laborious". Such meals are only planned on weekends.
- **Maximum number of laborious meals per weekend**. Are there none, just one, or even two laborious meals allowed per weekend?
- **Maximum popularity of a controversial meal**. Meals whose popularity does not exceed this value are considered "controversial".
- **Maximum number of controversial meals per planning period**. How many controversial meals may occur per planning period?
- **Planning period length.** It is determined automatically as the sum of all category repetitions. **Please do not change it manually.**
- Table **Frequency of categories**. How often shall which categories be planned? For example, we eat noodles or poultry roughly once a week, but fish only about every two weeks.

How do you decide on the right length of the planning period?

- When the planning period is too short, e.g. just one week, then you cannot plan any meal category less frequently than once a week. Such a planning would e.g. not be varied enough for us.
- When the planning period is too long, e.g. one year. Who wants to plan their meals a year in advance?

Therefore, we decided on a planning period of 28 days (4 weeks).

Caution: If your rules are too strict (or if there are too few meals in your own meal collection), then you can "corner" the automatic plan generation and effectively exclude some meals, e.g. laborious or unpopular ones.

**Use case: The meal category "Turbo".**

We have included some ultra-simple meals in our Turbo category. These are not full-fledged dinner meals, but rather simple breakfast-style dishes. We do not include this category in the list of our category frequencies on the worksheet Rules. Hence, during the automatic planning, these turbo meals are deliberately not planned at all. We only use them for manual planning whenever we want to "cheat" on a given day and want to plan an ultra-fast meal there for some reason.

**Worksheet Planning**

This is where the planning takes place.

- **Plan start**. The first day of the meal plan.
- **Today's date** button. It sets the start of the meal plan to today's date.
- **Plan length**. The length of the meal plan in days. Ideally, you should use the same length here as the length of the planning period (on the Rules worksheet). However, in principle you can also choose other lengths of your meal plan.
 If you generate a shorter plan than one planning period, then a full planning period will be generated internally and its unnecessary part will simply be thrown away. If you generate a plan longer than one planning period, then this will be composed of elementary building blocks of one planning period each.
- **Plan archived**. Has this plan already been archived? (0 - no, 1 - yes) The value is filled automatically and should not be changed manually.
- **Generate**. This button generates a meal plan with the desired start and length.

The meal plan generation considers the specified number of repetitions for each category. The categories are randomly mixed for every planning period. The distances between the meals of the same category are optimized. (The average spacing is determined by the number of repetitions within the planning period. The tolerated spacing values are average spacing +/- 1/3.) This ensures that the same meal category never lands on two consecutive days.

For each category, the meal with the most distant past serving date is selected. The information about the last serving date from the meal worksheet is used to achieve this. Therefore, you should ideally never delete last serving dates in the meal sheet. If there are several equally good candidate meals, then one of them is chosen randomly. The rules for the elaborate and controversial meals (see the sheet Rules) are taken into account as far as possible during the generation. The laborious and controversial meals are also color-highlighted accordingly in the meal plan.

Even if the automatic planning takes multiple aspects into account, it is still clearly inferior to an intelligent human planning. Therefore, you should check every automatically generated meal plan and improve it manually, e.g. by swapping and changing meals.

One example of the human superiority in the meal planning is the recognition of similarity of meals. Our generation algorithm currently only recognizes the similarity via the meal category. Other similarities are not taken into account. For example, a lentil soup can be generated right next day after cooked lentils, just because one of these meals belongs to the category "soup" and the other one to the category "vegetarian meal". Considering similarities in general is a challenging requirement because the similarities are multi-dimensional. Two meals can be similar in terms of an ingredient, their consistency or method of preparation (e.g. both deep-fried), their meat content, their taste, their cultural background (e.g. both Asian meals), or in terms of their other characteristics. Recognizing all of these similarity dimensions and simultaneously optimizing the distances for all of them would be a challenging optimization task.

Also personal considerations, such as "I have a long working day on the 17th, so I only want to plan a very simple meal there." can, of course, not be taken into account during automatic planning. All such cases need to be edited manually.

There are following possibilities for manual meal plan improvements:

- **Swap 2 meals**. This button can be used to swap two meals within the plan. Select two cells in two different rows within the current meal plan first, and then press this button. The two meals will be swapped.
- **Change meal…** You can select another meal of the same category via this button. Unfortunately, Excel does not allow any specific input dialogs. Hence, you can only select a new meal via a number in the input edit box. Effort and popularity are shown in the square brackets for each item, e.g. [e:2 p:3]. If known, also the nearest serving date is shown in the round parentheses. The nearest serving can be located in the past or in the future, e.g. (+10d). When determining the closest serving dates, both the dates from the current plan as well as the last serving dates from the meal worksheet are considered.
 As the width of this Excel input dialog cannot be modified, we have to abbreviate the textual information as much as possible. This is the reason for the minimalist and slightly cryptic format of the effort, popularity and date information. However, if our texts were longer, then it would lead to even more line breaks and, consequently, worse readability.
- **Change category…** This button can be used to select a different meal category. For each category, the closest serving date is given in the parentheses. The meal, which has not been served for the longest time, is initially selected for the new category. If this meal does not suit you, you can later change it using the "Change meal" button. Via the "Change category" function you can, for instance also select the "Turbo" meal category, which is never generated automatically in our example.
- **Delete meal**. You can use this button to delete one or more meals if you don't want to cook at all, e.g. because you're traveling or you want to eat out that day.
- The standard Excel functions " **Print**" or " **Save as**" – "Pdf". Once you are satisfied with your meal plan, you can print it or export it as a pdf. The print area of the worksheet is automatically set to the plan table.
- **Generate shopping lists**. Once you are satisfied with the meal plan, you can use this button to generate your weekly shopping lists. These can later be copied to an e-mail or some 3-rd party shopping list by using the standard copy & paste, or can also be printed out. The shopping lists are automatically deleted every time the meal plan changes, otherwise their content would become obsolete. If needed, they can be generated at any time via this button again.

Once the meal plan has been sufficiently edited, it can be archived. Afterwards, the next period can be planned.

- **Archive**. Once you are satisfied with your meal plan after the manual changes, you can archive it via this button. This archives the serving dates of all planned meals to the Meals worksheet. After that, this meal plan should not be changed on the Planning sheet. And exactly because you may want to change your meal plan spontaneously during its use (we do that for our plans regularly), it is better not archiving it until its planning period is really over. Hence, ideally do it just before you move on to the next planning period. As the archiving operation cannot be easily undone, it is recommended to save your Excel file before executing it. This will be also suggested to you by a corresponding message.
- **Next period**. After archiving, you can use this button to go to the next planning period. The new Plan Start = current Plan Start + current Plan Length. The planning cycle then begins again, i.e.: Generate  manually improve  Export and, if needed, also Generate shopping lists  Archive  Next period. And so forth, in a continuous cycle.

**Use case: "Experiment"**

If you wish to try out a whole new recipe from the Internet e.g. once a week, then you can add a new meal "Experiment" to your Meals worksheet. Assign it a unique category "Experiment" and set its desired frequency on the Rules worksheet. As a consequence, the placeholder meal "Experiment" will be generated in your meal plans with the specified frequency.

**Use case: "Refresh"**

When you generated a meal plan and you later edited some meals in your Meals worksheet (e.g. their categories, efforts, or other attributes), then your current meal plan may no longer be consistent afterwards. (Some meal in your plan may e.g. display the wrong effort or category.) Such an inconsistent meal plan can be updated via the Hotkey **Ctrl+Shift+R** (R like "Refresh"). As long as the meal names haven't been changed, this macro can fix all other data in the meal plan. If meal names have also been changed, then these must be corrected manually first, before the macro can be used.

**Language worksheet**

All user messages and UI elements of this Excel sheet are internationalized. On the worksheet Language you can select the UI language in the " **Current language"** field and then activate it with the " **Change language**" button.

You can also add an additional column for **a new language** in the table on this sheet, e.g. "SP" for Spanish. In this new column you can enter your own string translations and the sheet will support this new UI language.

New language columns should ideally be inserted in the penultimate position, i.e. to the left of the last language column. The new language will then be automatically offered in the language selection combo box. (Otherwise you would have to expand the data source for the combo box contents via the Excel function: "Data" -\> "Data verification".)

The strings in the table can include other strings via the \<StringId\> syntax and can also position parameter values via the # placeholder. The parameter contents are determined by the VBA code which uses the respective string.

The string **Id** -s in the column A must not be changed.

You can invoke the about box containing the information about this tool, its author and its license terms via the " **About…**" button.

**Known limitations**

- The meal plan targets the use case with one meal per day. This might be the lunch for some families and the dinner for others. A version with 2, 3 or even up to 5 meals per day would also be thinkable. However, this is not our use case, so I currently do not intend to implement this.
- There are many thinkable strategies for a meal plan generation, such as weekday-related meal categories, e.g. "always cooking fish on Fridays", "always cooking meat on Sundays", etc. One could implement several such strategies and select the active one via a new combo box on the Planning sheet before pressing the Generate button. However, this is not our use case and I currently do not intend to implement this.

If anyone wishes such improvements, feel free to extend the macros. The project is freeware & open source.

**Technical Notes**

- The worksheets are _not_ identified by their titles (otherwise the internationalization would not work). They have fixed variable names that can only be changed via the macro editor. Therefore, it would not work if you e.g. renamed the current "Meals" tab to "Meals\_old " and then created your own new "Meals" sheet. The tool would continue using the dishes from the "Food\_old" anyway. Hence, if you want to import an alternate food collection, you have to copy its content to the original worksheet Meals.
- All input fields are expected on fixed (row & column) coordinates. Therefore, you cannot insert new rows or columns at an arbitrary place of the worksheets. However, adding new data to the very bottom or the far right of the worksheets should not be a problem.
- There are internal constants in the VBA code, which limit the maximum number of meal categories, the maximum length of the planning period, and many more parameters. If these limits are exceeded during the use, then an assertion can be violated and the VBA debugger will be invoked. If you have a basic VBA knowledge, then you can easily increase such constants in the code.
- Button flickering. When some macros modify e.g. the bold text formatting or the frame formatting of the cells, then all graphic objects (i.e. the buttons) on the worksheet flicker once. This happens in older Excel versions. It is an Excel bug and it cannot be suppressed. However, this problem is no longer observed in Excel 365.
