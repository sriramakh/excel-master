"""Marketing dataset: Campaign performance, lead funnel, channel analytics."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .base import BaseGenerator, rng_choice, rng_uniform, rng_integers, rng_normal, date_range, make_ids

CHANNELS = ["Paid Search", "Social Media", "Email", "Content/SEO", "Display Ads",
            "Events/Webinars", "Partner", "Direct Mail", "Podcast", "Influencer"]
CAMPAIGN_TYPES = ["Brand Awareness", "Lead Generation", "Retargeting", "Product Launch",
                   "Customer Retention", "Upsell/Cross-sell", "Account Based Marketing"]
PLATFORMS = {
    "Paid Search": ["Google Ads", "Microsoft Bing"],
    "Social Media": ["LinkedIn", "Meta (Facebook/IG)", "Twitter/X", "YouTube"],
    "Email": ["HubSpot", "Marketo", "Salesforce Marketing Cloud"],
    "Content/SEO": ["Blog", "Landing Pages", "Video", "Webinar On-Demand"],
    "Display Ads": ["Google Display", "The Trade Desk", "AppNexus"],
    "Events/Webinars": ["Dreamforce", "Virtual Summit", "Regional Events"],
    "Partner": ["Agency Partners", "Resellers", "Tech Partners"],
    "Direct Mail": ["Physical Mailers", "Gifting"],
    "Podcast": ["Spotify", "Apple Podcasts", "YouTube"],
    "Influencer": ["Industry Analysts", "Micro-influencers"],
}
INDUSTRIES_TARGET = ["Technology", "Finance", "Healthcare", "Retail", "Manufacturing"]
GEO_TARGETS = ["North America", "Europe", "APAC", "Global"]
FUNNEL_STAGES = ["Awareness", "Interest", "Consideration", "Intent", "Evaluation", "Purchase"]
LEAD_STATUSES = ["MQL", "SAL", "SQL", "Opportunity", "Won", "Lost", "Nurture", "Disqualified"]
CONTENT_TYPES = ["Blog Post", "Whitepaper", "Case Study", "Webinar", "Demo", "Video",
                  "Infographic", "Report", "Newsletter", "Podcast Episode"]


def _campaigns(n: int = 2000) -> pd.DataFrame:
    channels = rng_choice(CHANNELS, n)
    campaign_types = rng_choice(CAMPAIGN_TYPES, n)
    start_dates = date_range("2023-01-01", "2024-12-01", n)
    durations = rng_integers(7, 90, n)
    end_dates = pd.to_datetime(start_dates) + pd.to_timedelta(durations, unit="D")

    budget = rng_uniform(2000, 150000, n).round(2)
    spend = (budget * rng_uniform(0.70, 1.05, n)).round(2)
    impressions = rng_integers(5000, 5000000, n)
    clicks = (impressions * rng_uniform(0.005, 0.08, n)).astype(int)
    ctr = (clicks / impressions * 100).round(2)
    cpc = np.where(clicks > 0, spend / clicks, 0).round(2)
    leads = (clicks * rng_uniform(0.02, 0.20, n)).astype(int)
    mql = (leads * rng_uniform(0.3, 0.7, n)).astype(int)
    sql = (mql * rng_uniform(0.2, 0.5, n)).astype(int)
    opportunities = (sql * rng_uniform(0.3, 0.8, n)).astype(int)
    deals_closed = (opportunities * rng_uniform(0.15, 0.45, n)).astype(int)
    avg_deal = rng_uniform(5000, 85000, n)
    revenue_influenced = (deals_closed * avg_deal).round(2)
    roi = np.where(spend > 0, (revenue_influenced - spend) / spend * 100, 0).round(1)
    cpl = np.where(leads > 0, spend / leads, 0).round(2)
    cac = np.where(deals_closed > 0, spend / deals_closed, 0).round(2)

    platforms = [rng_choice(PLATFORMS[c], 1)[0] for c in channels]

    df = pd.DataFrame({
        "campaign_id": make_ids("CMP", 10001, n),
        "campaign_name": [f"{ct} - {ch} - {rng_integers(2023, 2024, 1)[0]}-Q{rng_integers(1, 4, 1)[0]}"
                          for ct, ch in zip(campaign_types, channels)],
        "channel": channels,
        "platform": platforms,
        "campaign_type": campaign_types,
        "start_date": pd.to_datetime(start_dates).date,
        "end_date": end_dates.date,
        "duration_days": durations,
        "target_industry": rng_choice(INDUSTRIES_TARGET, n),
        "geo_target": rng_choice(GEO_TARGETS, n),
        "budget_usd": budget,
        "spend_usd": spend,
        "budget_utilization_pct": (spend / budget * 100).round(1),
        "impressions": impressions,
        "clicks": clicks,
        "ctr_pct": ctr,
        "cpc_usd": cpc,
        "leads_generated": leads,
        "mql_count": mql,
        "sql_count": sql,
        "opportunities_created": opportunities,
        "deals_closed": deals_closed,
        "revenue_influenced_usd": revenue_influenced,
        "avg_deal_size_usd": avg_deal.round(2),
        "cost_per_lead_usd": cpl,
        "cost_per_acquisition_usd": cac,
        "roi_pct": roi,
        "roas": np.where(spend > 0, revenue_influenced / spend, 0).round(2),
        "conversion_rate_pct": np.where(leads > 0, deals_closed / leads * 100, 0).round(2),
        "lead_to_mql_rate": np.where(leads > 0, mql / leads * 100, 0).round(1),
        "mql_to_sql_rate": np.where(mql > 0, sql / mql * 100, 0).round(1),
        "quarter": pd.to_datetime(start_dates).quarter,
        "month": pd.to_datetime(start_dates).month,
        "year": pd.to_datetime(start_dates).year,
        "status": rng_choice(["Completed", "Active", "Paused", "Planned"], n,
                              p=[0.55, 0.25, 0.12, 0.08]),
        "ab_test_variant": rng_choice(["A", "B", "Control", "N/A"], n, p=[0.15, 0.15, 0.10, 0.60]),
    })
    return df.sort_values("start_date").reset_index(drop=True)


def _web_analytics(months: int = 24) -> pd.DataFrame:
    """Monthly web analytics data."""
    dates = pd.date_range("2023-01-01", periods=months, freq="MS")
    base_sessions = 45000
    rows = []
    for i, d in enumerate(dates):
        trend = 1 + 0.015 * i
        sessions = int(base_sessions * trend * rng_uniform(0.9, 1.1, 1)[0])
        users = int(sessions * rng_uniform(0.7, 0.85, 1)[0])
        pageviews = int(sessions * rng_uniform(2.5, 4.5, 1)[0])
        bounce_rate = rng_uniform(38, 65, 1)[0].round(1)
        avg_session = rng_uniform(1.5, 5.5, 1)[0].round(1)
        conversions = int(sessions * rng_uniform(0.02, 0.06, 1)[0])
        rows.append({
            "month": d.strftime("%Y-%m"),
            "year": d.year,
            "month_num": d.month,
            "quarter": f"Q{(d.month-1)//3+1}",
            "total_sessions": sessions,
            "unique_users": users,
            "new_users": int(users * rng_uniform(0.55, 0.75, 1)[0]),
            "returning_users": int(users * rng_uniform(0.25, 0.45, 1)[0]),
            "pageviews": pageviews,
            "pages_per_session": round(pageviews / sessions, 1),
            "bounce_rate_pct": bounce_rate,
            "avg_session_duration_min": avg_session,
            "organic_sessions": int(sessions * rng_uniform(0.30, 0.45, 1)[0]),
            "paid_sessions": int(sessions * rng_uniform(0.20, 0.35, 1)[0]),
            "social_sessions": int(sessions * rng_uniform(0.10, 0.20, 1)[0]),
            "direct_sessions": int(sessions * rng_uniform(0.10, 0.20, 1)[0]),
            "referral_sessions": int(sessions * rng_uniform(0.05, 0.10, 1)[0]),
            "form_conversions": conversions,
            "conversion_rate_pct": round(conversions / sessions * 100, 2),
            "demo_requests": int(conversions * rng_uniform(0.3, 0.5, 1)[0]),
            "free_trials": int(conversions * rng_uniform(0.2, 0.4, 1)[0]),
        })
    return pd.DataFrame(rows)


def _content_performance(n: int = 300) -> pd.DataFrame:
    content_types = rng_choice(CONTENT_TYPES, n)
    publish_dates = date_range("2023-01-01", "2024-12-01", n)
    views = rng_integers(100, 50000, n)
    downloads = np.where(rng_choice(["download", "view"], n, p=[0.4, 0.6]) == "download",
                          rng_integers(10, 5000, n), 0)
    leads = (downloads * rng_uniform(0.05, 0.25, n)).astype(int)
    df = pd.DataFrame({
        "content_id": make_ids("CNT", 1001, n),
        "title": [f"{ct} #{rng_integers(100, 999, 1)[0]}: Topic About Industry" for ct in content_types],
        "content_type": content_types,
        "channel": rng_choice(CHANNELS[:6], n),
        "publish_date": pd.to_datetime(publish_dates).date,
        "quarter": pd.to_datetime(publish_dates).quarter,
        "views": views,
        "unique_visitors": (views * rng_uniform(0.6, 0.9, 1)[0] * np.ones(n)).astype(int),
        "downloads": downloads,
        "shares": rng_integers(0, 500, n),
        "comments": rng_integers(0, 50, n),
        "avg_time_on_page_sec": rng_integers(30, 480, n),
        "bounce_rate_pct": rng_uniform(25, 80, n).round(1),
        "leads_generated": leads,
        "influenced_pipeline_usd": (leads * rng_uniform(5000, 30000, n)).round(2),
        "seo_rank": rng_choice([*range(1, 11), None], n,
                               p=[0.08, 0.08, 0.08, 0.07, 0.07, 0.06, 0.06, 0.06, 0.05, 0.05, 0.34]),
        "backlinks": rng_integers(0, 200, n),
        "engagement_score": rng_uniform(10, 100, n).round(1),
    })
    return df


class MarketingGenerator(BaseGenerator):
    name = "marketing"
    industry = "Marketing"
    description = "Campaign performance, web analytics, and content effectiveness data"

    def generate(self) -> dict[str, pd.DataFrame]:
        print("  Generating Campaign Data (2,000 rows)...")
        campaigns = _campaigns(2000)
        print("  Generating Web Analytics (24 months)...")
        web = _web_analytics(24)
        print("  Generating Content Performance (300 rows)...")
        content = _content_performance(300)

        # Channel summary
        channel_summary = campaigns.groupby("channel").agg(
            campaigns_count=("campaign_id", "count"),
            total_spend=("spend_usd", "sum"),
            total_leads=("leads_generated", "sum"),
            total_revenue_influenced=("revenue_influenced_usd", "sum"),
            avg_roi=("roi_pct", "mean"),
            avg_cpl=("cost_per_lead_usd", "mean"),
            avg_conversion=("conversion_rate_pct", "mean"),
        ).round(2).reset_index()
        channel_summary["roas"] = (channel_summary["total_revenue_influenced"] /
                                    channel_summary["total_spend"]).round(2)

        return {
            "Campaigns": campaigns,
            "Web_Analytics": web,
            "Content_Performance": content,
            "Channel_Summary": channel_summary,
        }
