# Spec Formatter: hosted files

This repository only hosts the static web files for the **Spec Formatter** Microsoft Word add-in
(Systems West Engineers, internal tool). It contains no source code.

- **Where the files are.** GitHub Pages serves the **`gh-pages`** branch at
  `https://dev-swe.github.io/SpecRevFormat/`. This `main` branch holds only this README.
- **Generated, not edited.** The `gh-pages` files are built and published automatically from a
  private source repository. Do not edit them by hand; changes are overwritten on the next publish.
- **Not a website to use directly.** The add-in runs inside Word. Its home page only explains that.
- **Do not rename or delete this repository.** The add-in's deployed manifest points at the URL
  above, so changing it breaks the add-in for every user until a new package is deployed.
- **Internal use only.** No license is granted for reuse.

Questions: contact Systems West Engineers.
