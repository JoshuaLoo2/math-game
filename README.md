# Math Game Question Schema

The Google Sheets question bank now supports three question types for both `getQuestions()` and `getNotesQuiz()`.

Recommended setup: use three separate worksheets inside the same spreadsheet.

- Worksheet `multipleChoice`: store all multiple-choice questions here.
- Worksheet `fillBlank`: store all fill-in-the-blank questions here.
- Worksheet `matching`: store all matching questions here.

When you use this worksheet-per-type setup, the backend will infer `questionType` from the worksheet name, so the `questionType` column is optional.

Recommended columns:

- `question`
- `questionType` (optional when the worksheet name already indicates the type)
- `answer`
- `acceptedAnswers` (optional; line-separated alternative correct answers)
- `optionA`
- `optionB`
- `optionC`
- `optionD`
- `matchingLeft`
- `matchingRight`
- `explanation`
- `imageUrl`
- `points` (notes quiz only)

Examples:

## Fill Blank

- Worksheet: `fillBlank`
- `questionType`: `fillBlank`
- `question`: `解方程 $3x + 7 = 22$，$x =$ ____`
- `answer`: `5`

## Matching

- Worksheet: `matching`
- `questionType`: `matching`
- `question`: `把算式和結果配對`
- `matchingLeft`:
	- `6 × 7`
	- `9 + 8`
	- `15 - 6`
- `matchingRight`:
	- `42`
	- `17`
	- `9`

## Paste-ready Google Sheet Templates

Copy the block for each worksheet and paste it into cell `A1` in Google Sheets.

### Challenge Question Bank: `getQuestions()`

Worksheet: `multipleChoice`

```tsv
theme	difficulty	grade	question	answer	optionA	optionB	optionC	optionD	explanation	imageUrl
代數	easy	中一	如果 $3x+7=22$，那麼 $x$ 等於多少？	5	3	4	5	6	$3x=15$，所以 $x=5$。	
幾何	medium	中一	一個正方形邊長是 8 cm，面積是多少？	64 cm²	48 cm²	56 cm²	64 cm²	72 cm²	正方形面積 = 邊長 × 邊長 = $8 \times 8 = 64$ cm²。	
```

Worksheet: `fillBlank`

```tsv
    theme	difficulty	grade	question	answer	acceptedAnswers	explanation	imageUrl
代數	easy	中一	解方程 $3x + 7 = 22$，$x =$ ____	5		$3x=15$，所以 $x=5$。	
百分數	medium	中一	把 $0.25$ 化成百分數：____ %	25	25%	$0.25 \times 100\% = 25\%$。	
```

Worksheet: `matching`

```tsv
theme	difficulty	grade	question	answer	matchingLeft	matchingRight	explanation	imageUrl
四則運算	easy	中一	把算式和答案配對	6×7->42 | 9+8->17 | 15-6->9	6 × 7
9 + 8
15 - 6	42
17
9	每條算式都要對應正確結果。	
分數	medium	中一	把分數和小數配對	1/2->0.5 | 1/4->0.25 | 3/4->0.75	1/2
1/4
3/4	0.5
0.25
0.75	先把分數轉成小數再配對。	
```

### Notes Quiz Bank: `getNotesQuiz()`

Worksheet: `multipleChoice`

```tsv
topic	grade	section	question	answer	points	optionA	optionB	optionC	optionD	explanation	imageUrl
代數	中一	一元一次方程	若 $x + 5 = 12$，那麼 $x$ 等於多少？	7	10	5	6	7	8	把 5 移到右邊：$x=12-5=7$。	
百分數	中一	百分數概念	$40\%$ 用小數表示是甚麼？	0.4	10	0.04	0.4	4	40	百分數除以 100，所以 $40\%=0.4$。	
```

Worksheet: `fillBlank`

```tsv
topic	grade	section	question	answer	acceptedAnswers	points	explanation	imageUrl
    代數	中一	一元一次方程	若 $2x = 18$，那麼 $x =$ ____	9		10	兩邊同除以 2，所以 $x=9$。	
百分數	中一	百分數概念	$75\%$ 用小數表示是 ____	0.75	.75	10	百分數除以 100，所以 $75\%=0.75$。	
```

Worksheet: `matching`

```tsv
topic	grade	section	question	answer	points	matchingLeft	matchingRight	explanation	imageUrl
代數	中一	一元一次方程	把方程和解配對	$x+3=8$->5 | $2x=14$->7 | $x-4=6$->10	10	$x+3=8$
$2x=14$
$x-4=6$	5
7
10	每條方程各自解出 $x$ 的值。	
分數	中一	分數與小數	把分數和小數配對	1/5->0.2 | 2/5->0.4 | 4/5->0.8	10	1/5
2/5
4/5	0.2
0.4
0.8	把每個分數化成小數後再配對。	
```

### Notes

- The worksheet names should be exactly `multipleChoice`, `fillBlank`, and `matching`.
- `answer` is still required in all worksheets because the backend expects that column.
- In `matching`, put the left items and right items in the same row, using line breaks inside each cell.
- `acceptedAnswers` can contain multiple correct answers, separated by line breaks.
