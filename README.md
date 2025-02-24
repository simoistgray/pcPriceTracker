# PC Part Price Tracker

## Overview
The PC Part Price Tracker is a Python application built for tracking the price of specific PC parts over time. The application allows users create a build on pcpartpicker and then track the prices. The application can run automatically incrementally using crontab. Users will get email notifications when a part hits a new sale or dips below a certain price.

## Features
- Track prices of individual PC parts from PCPartPicker.
- Set a target price for each part to receive alerts when the price drops below it.
- Automatically fetch updated prices at scheduled intervals using cron jobs.
- Send email notifications when a tracked part is on sale or meets the user's target price.
- Log price changes over time for historical tracking and analysis.
- Simple and customizable configuration file for managing tracked parts and notification settings.

## Future Improvements
- Add support for multiple online retailers beyond PCPartPicker.
- Implement a web interface for easier configuration and tracking.
- Add a database to store price history and generate trend reports.
- Support for Discord notifications in addition to email.

## License
This project is open-source under the MIT License.

## Author
Simon Gray
