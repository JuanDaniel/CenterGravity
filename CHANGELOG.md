# Changelog

All notable changes to **Center Gravity** are documented here.
The format is based on [Keep a Changelog](https://keepachangelog.com/).

## [2.0.0] - 2026-09-08

### Compatibility

- Added support for **Revit 2024, 2025, 2026 and 2027**.
- The add-in now ships as a single Autodesk **ApplicationPlugins bundle**
  (`CenterGravity.bundle`); one installer covers every supported release and
  Revit loads the matching build automatically.

### Added

- **Dockable panel.** The floating dialog is replaced by a panel that docks
  like *Properties*, updates live as you change the selection, follows the
  Revit light/dark theme and can be resized. Selected elements are listed with
  their individual volume and mass; elements without solid geometry are shown
  dimmed and reported in a status line.
- **Centre of gravity by mass.** Choose *Volume* or *Mass* weighting. Mass uses
  each material's structural density; a **default density** can be entered for
  materials that carry none. The total mass is displayed alongside the volume.
- **Assemblies and groups.** Selecting an assembly or a model group expands it
  to its member elements automatically (nested groups included) - no need to
  edit the assembly or explode the group.
- **Schedulable / taggable marker.** The centre-of-gravity marker now carries
  shared parameters: coordinates, volume, weight, lift name, reference point
  and date. A **Schedule** button builds a *Center of Gravity* schedule of all
  markers in one click.
- **Reference point.** Report coordinates from the internal origin, the project
  base point or the survey point.
- **Lift name.** A free-text label written onto each marker, so several lifts
  can be told apart in a schedule.
- **Preliminary rigging checks.** Pick 2 or 4 lift points to get the vertical
  load share per leg and its percentage of the total weight, a check that the
  centre of gravity projects inside the lift-point polygon, and an uplift
  warning. Enter a hook height above the centre of gravity to also get the
  sling angle and tension per leg and the tilt the load would take if rigged
  symmetrically. This is a preliminary aid, **not a certified lift plan**.
- **Export.** Copy the coordinate string, or export the element list, totals
  and rigging results to **CSV**.

### Changed

- **The marker is now opt-in.** The centre-of-gravity point is placed only when
  you press **Place marker**; changing the selection no longer drops geometry
  into the model. **Clear** removes it.
- Selection changes are detected through the Revit selection event instead of
  the ribbon, which is more reliable and independent of the interface language.
- The tool computes from whatever is already selected as soon as the panel is
  opened.

### Fixed

- The dockable panel could come up empty and unresponsive to selection changes
  (most visible on Revit 2027): its content was built during Revit's very
  early `ApplicationInitialized` event, which is not guaranteed to have
  finished - or to have Revit's WPF hosting ready - before the panel can be
  shown, and a dockable pane's content is only ever built once per session.
  It is now built the first time the command actually runs, which is always
  after Revit is fully interactive.
- A selection with no solid geometry no longer produces an empty result or an
  error; the panel reports which elements were left out.
- Corrected the weighting of elements built from more than one solid.
- Re-selecting the same element no longer raises an internal error.
- Coordinate and volume values now follow the project's unit and rounding
  settings.

## [1.0.6] - 2025

- Multi-target build for Revit 2021-2025.
- Inno Setup installer.
