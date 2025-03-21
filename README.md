# sql-historical-IT-chargebacks

This script was created to serve as a master query used for financial and resource management reports. This histrical report is cut off at the time of migration to a web application instead of on-premise.

The source does not have a dynamic table for resource rates and cost centers. Therefore the query refers to a custom managed look up table to determine the rate and cost center for the recorded date range. This query includes accounting code calculation as well as other calculated fields.
