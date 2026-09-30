# Verification report (照合レポート)

- ground truth: 1300 rows, seed 42
- output rows found: 1300
- rows missing from the output: 0

## 1. Planted errors: planted / detected as expected / rate (仕込んだ数 / 検出数 / 検出率)

| kind | planted | detected | rate |
| --- | ---: | ---: | ---: |
| duplicate_exact | 40 | 40 | 100.0% |
| duplicate_email | 30 | 30 | 100.0% |
| possible_duplicate | 15 | 15 | 100.0% |
| invalid_email | 50 | 50 | 100.0% |
| email_domain_typo | 10 | 10 | 100.0% |
| ambiguous_date | 20 | 20 | 100.0% |
| invalid_date | 5 | 5 | 100.0% |
| invalid_phone | 15 | 15 | 100.0% |
| missing_name | 10 | 10 | 100.0% |
| no_contact | 10 | 10 | 100.0% |
| phone_leading_zero | 10 | 10 | 100.0% |
| soft_missing | 100 | 100 | 100.0% |
| phone_format_unverified | 5 | 5 | 100.0% |

- `duplicate_of` correct: 85 / 85
- `suggestion` correct (typo / ambiguous date): 40 / 40

## 2. Rows with unexpected category (期待と違う区分): 0 rows

- none

## 3. Normalization accuracy on clean rows (正規化精度)

- expected-clean rows that were output as clean: 1085 / 1085
- all 6 columns equal to canonical: 1085 / 1085 = 100.00%


## 4. False positives: expected-clean rows flagged invalid / needs_review / duplicate (誤検出): 0


## 5. Targets (目標)

- PASS: duplicate detection >= 95%
- PASS: invalid email detection >= 95%
- PASS: false positives ~ 0
- PASS: normalization accuracy >= 98%
- PASS: every planted kind detected as expected (no unexpected category)
