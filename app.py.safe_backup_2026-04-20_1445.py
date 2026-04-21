    (
        use_vat_budget_metrics,
        use_ak_budget_metrics,
        ak_mode,
        ak_fixed_month_wo_vat,
        ak_fixed_percent,
        ak_rules_df,
    ) = render_vat_ak_section(
        metric_mode=metric_mode,
        is_real_estate_preset=is_real_estate_preset,
        use_vat_budget_metrics=use_vat_budget_metrics,
        use_ak_budget_metrics=use_ak_budget_metrics,
        ak_mode=ak_mode,
        ak_fixed_month_wo_vat=ak_fixed_month_wo_vat,
        ak_fixed_percent=ak_fixed_percent,
        ak_rules_df=ak_rules_df,
    )
