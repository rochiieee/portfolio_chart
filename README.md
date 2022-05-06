# portfolio_chart

As owner, this repo contains the original source code to the portfolio growth chart google apps script function(s) presented on https://www.youtube.com/watch?v=UCWfdW6hcVA&t=1731s

For the script to run, timer's must be implemented on the google apps script page, and be set proportionally to the lastrow variable under each unique function. I.E for a 1-year timeframe, if you intend to record the value every 8 hours, 3*365 must become the value of lastRow. Thus, for all unique possible timeframe-frequency possibilities, timeframe (in years) * 365 * (24/update (hours) ).

Atop such, your total_value variable (equals summated value of investments), must be placed in cell A1.


