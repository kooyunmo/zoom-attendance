# zoom-attendance

## Usage

``` bash
python attendance.py \
  -i input.xlsx \
  -t 120 \              # lecture duration (min)
  -s 18:30 \            # lecture start time (HH:MM)
  -l 10 \               # late threshold (min)
  -o output.xlsx
```

**Options**

- `-i` or `--input-file`: input .xlsx file name
- `-t` or `--min-time`: the minimum time to admit attendance
- `-s` or `--start-time`: the start time of the lecture (HH:MM)
- `-l` or `--late-time`: students participating later than `late-time` are regarded as 'late'.
- `-o` or `--output-file`: output .xlsx file name
