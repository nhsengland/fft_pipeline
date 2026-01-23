# Ward Ranking Logic Discrepancy - Root Cause Analysis

## Issue Summary

Validation failures showing 24 differences in Ward sheet suppression between Python pipeline output and VBA ground truth.

## Root Cause Identified

**Tie-breaking Order Mismatch** for wards with equal Total Responses:

### VBA Logic:
- Primary: Total Responses (ascending)
- Secondary: **First Specialty** (alphabetical)
- Tertiary: **Second Specialty** (alphabetical)
- Quaternary: **Ward Name** (alphabetical)

### Original Python Logic:
```python
sorted_indices = df_temp.sort_values(
    ["Total Responses", "Ward_Name", "_spec1_text", "_spec2_text"]
).index
```

### Corrected Python Logic:
```python
sorted_indices = df_temp.sort_values(
    ["Total Responses", "_spec1_text", "_spec2_text", "Ward_Name"]
).index
```

## Evidence from Specific Cases

**Site RWD|RWDLA** (both wards have 6 responses):
- Ward "7A": Specialty `800 - CLINICAL ONCOLOGY`
- Ward "Ward 1": Specialty `326 - ACUTE INTERNAL MEDICINE`
- VBA ranks by specialty: "326..." < "800..." → Ward 1 gets rank 2 (suppressed)
- Original Python ranked by ward name: "7A" < "Ward 1" → 7A got rank 2 (wrong)

**Site RKE|RKEQ4** (both wards have 6 responses):
- Ward "Cloudesley": Specialty `430 - GERIATRIC MEDICINE`
- Ward "Victoria": Specialty `301 - GASTROENTEROLOGY`
- VBA ranks by specialty: "301..." < "430..." → Victoria gets rank 2 (suppressed)
- Original Python ranked by ward name: "Cloudesley" < "Victoria" → Cloudesley got rank 2 (wrong)

## Current Status (Post-Fix)

✅ **Specialty-first fix CONFIRMED working**
- Original failing wards (RWD|RWDLA, RKE|RKEQ4) completely resolved
- Oct-25: Still 24 differences but **different wards** (RQ3|RQ301, R1F|R1F01)
- Validates that tie-breaking fix addressed specific VBA logic mismatch

⚠️ **Additional ranking issues identified**
- Jul-25: 132 differences unchanged (RTD|RTD06, RGR|RGR50, RWD|RWDDA sites)
- Different root cause from specialty tie-breaking issue
- Suggests multiple distinct VBA ranking logic variations

🚨 **Jun-25 systemic problems persist**
- ICB level: 86 percentage precision differences (`0.9501764583765829` vs `0.9502`)
- Site level: 74 suppression logic mismatches (not ranking-related)
- Ward level: 284 differences (broader than ranking issues)

## Multiple Issues Framework

**Issue 1: Ward tie-breaking order** ✅ FIXED
- Specialty-first vs ward-name-first sorting
- Resolved for specific ward pairs

**Issue 2: Alternative VBA ranking logic** 🔍 INVESTIGATING
- Different sites may use different ranking criteria
- Jul-25 sites unaffected by specialty fix

**Issue 3: Percentage calculation differences** ✅ FIXED
- **Root cause**: Excel formatting `"0%"` rounded display values
- **Solution**: Changed to `"0.0000%"` + tolerance adjustment to `1e-5`
- **Result**: Jun-25 ICB/Trusts/Sites now validate perfectly (0 differences)

**Issue 4: Site-level suppression logic** ✅ RESOLVED
- **Root cause**: Same as Issue 3 (formatting, not logic)
- **Confirmed**: All site-level differences were display formatting artifacts

## Fix Results Summary

**Jun-25 Validation - Before vs After all fixes:**
- ICB: 86 → ✅ 0 differences
- Trusts: 2 → ✅ 0 differences
- Sites: 74 → ✅ 0 differences
- Wards: 284 → 36 differences (87% improvement)

**Remaining ward differences**: True ranking logic issues (not formatting)