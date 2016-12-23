- Start Date: 2016-12-15
- RFC PR: https://github.com/VBA-tools/VBA-Web/pull/267
- VBA-Web Issue: (leave this empty)

# Summary

Great care has been taken to ensure that all of VBA-Web works across Windows, Mac, and Office environments. This will continue to be a foundational principle of VBA-Web, but in the pursuit of new and platform-specific functionality, the use of "feature flags" is proposed. For a default installation, only core, cross-platform functionality will be enabled, but feature flags can be enabled to add features that the user determines can be supported for their application.

# Motivation

VBA-Web has a strong, stable core that is supported across all platforms, but it has been difficult to add some features that may have limited cross-platform support or are too intensive/experimental to add cross-platform support initially. Feature flags will allow development work to continue on VBA-Web with features that are planned for future cross-platform compatibility to be added with limited platform/application support initially. The goal is that the external API of VBA-Web should always be designed in a cross-platform manner, but feature flags allow the implementation to be limited to a single platform/application.

# Detailed design

Feature flags will be compiler constants that are `False` by default. The should start with "Enable" and then the feature name. Compatibility details should not be included in the feature name as this may change during development. Each flag should be isolated from all other feature flags. Behavior that depends on more than a single feature flag should be strongly discouraged and turning on one feature should not require turning on/off another. Generally, public functionality related to the feature should be put _inside_ compiler directives so that using disabled features causes compilation errors (rather than runtime errors).

```vb
' WebHelpers.bas

' <- Define feature flags and allow user override
''
' Feature description...
'
' @feature A
' @compatibility
'   Platforms: Windows
'   Applications: Excel
''
#Const EnableFeatureA = False

''
' Feature description...
'
' @feature B
' @compatibility
'   Platforms: Windows and Mac
'   Applications: Outlook
''
#Const EnableFeatureB = False

' -> Use feature flags with compiler directives
#If EnableFeatureA Then
Public Sub DoA()

End Sub
#End If
```

# How We Teach This

This should be strongly documented in the readme and docs, with up-to-date compatibility for each feature. These flagged features have the goal of becoming core, cross-platform features, so they will be supported and receive bug fixes with the primary project. There should be great care taken in selecting flagged features for a project and the implications for lock-in and other limitations that may arise for code-sharing and examples that use VBA-Web with a subset of feature flags enabled.

It is important that the installer at a minimum maintains feature flag selections during upgrades and may need to include selections for new installations.

# Drawbacks

Great care should be taken to ensure that all features added via feature flags have the potential of cross-platform support in the future. There may be feasibility issues, but fundamental functionality that will __never__ inherently work on another platform should be avoided as this would introduce a permanent divergence of the VBA-Web audience.

# Alternatives

Currently, `WebAsyncWrapper` is a separate class that is only compatible with Windows and Excel. Wrapping existing functionality has worked well for this specific case, but there are others that may need to be more integrated. Due to the cross-platform nature of most of these compatibility issues, simple options/arguments are not viable as some code may introduce compiler errors if they were not isolated with compiler directives.

# Unresolved questions

None.
