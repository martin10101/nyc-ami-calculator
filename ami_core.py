# ... (keep previous code, add auto prefs logic in generate_scenarios)
def generate_scenarios(df: pd.DataFrame, required_sf: float, prefs: Dict) -> Dict:
    df = normalize_df(df)
    valid = validate_selection(df, required_sf)
    if not valid["valid"]:
        raise ValueError(valid["message"])
    aff = df[df["AMI"].notna()].copy()
    # Auto-determine if 40% required
    prefs["required_40_pct"] = 0.20 if required_sf > 10000 else 0
    # Bake floor prefs: Sort aff by floor ascending for 40% assignment priority
    aff = aff.sort_values(by="FLOOR" if "FLOOR" in aff else aff.index)  # Low floors first for 40%
    # ... (rest same - MILP handles low floor pref via revenue weight negative for low floors on 40%)
    # In calculate_revenue_weight: Adjust for rules - low floors/small negative for high revenue, so high bands on high floors
    def calculate_revenue_weight(row):
        return (row.get("FLOOR", 1) * 0.3) + (row.get("BED", 0) * 0.4) + (row.get("NET SF", 500) / 1000 * 0.2) + (row.get("BALCONY", 0) * 0.1)  # High = high value

    # In MILP: Prefer high bands on high revenue (upper/large)
    obj = wavg * 100 - (0.60 - wavg) * 50 + lpSum(lpSum(assign[i][j] * bands[j] * revenue[i] for j in range(len(bands))) for i in range(n)) / n
    # ... (rest same)
