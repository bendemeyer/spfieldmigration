## About this project

This is a console app designed for migrating data between columns and column types within a list. Especially useful if you're transitioning to a new column type but don't want to lose the data currently housed in the old column.

This is a console app designed for migrating data between columns and column types within a list. Especially useful if you're transitioning to a new column type but don't want to lose the data currently housed in the old column.

One important feature of this app is how it handles minor versions. In order to preserve both the current minor version and the most recent major version, the app publishes a new version on top (based on the most recent major version), migrates the data there, then checks in another version on top of that (based on the most recent minor version) and migrates the data there as well. This is a must-have feature for publishing sites that wish to keep their public facing content up-to-date without losing any outstanding edits.

To run this console app, you must have appropriate permissions on both the SharePoint site collection you're targeting and the SQL content database it uses.

Outstanding issues: while most list items will not have their "modified" information changed, any items that have minor versions on top of the most recent major version will have their modified information update, and will show as modified by the user who runs the app.
