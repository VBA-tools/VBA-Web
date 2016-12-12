# VBA-Web RFCs

The RFC (request for comments) process is designed to open up the planning of future changes to VBA-Web to the community. Modeled after [Rust's](https://github.com/rust-lang/rfcs), [Ember's](https://github.com/emberjs/rfcs), and other community RFC projects, all substantial changes to VBA-Web (e.g. adding, changing, or removing major features) will go through a public design process where community members can provide feedback, suggest changes, and propose new features altogether.

## Audience

In addition to being open to community review and comment, community members are encouraged to open RFCs with changes they'd like to see in VBA-Web.

## Format

See `0000-template.md` for details, but generally, RFCs are design documents so they are more concerned with the overall API design of VBA-Web than the actual implementation details.

## The RFC Process

1. Fork the VBA-Web repo [https://github.com/VBA-tools/VBA-Web](https://github.com/VBA-tools/VBA-Web)
2. Copy rfcs/0000-template.md to rfcs/0000-my-feature.md (where 'my-feature' is descriptive. don't assign an RFC number yet)
3. Fill in the RFC. Put care into the details: __RFCs that do not present convincing motivation, demonstrate understanding of the impact of the design, or are disingenuous about the drawbacks or alternatives tend to be poorly-received__
4. Submit a pull request. As a pull request the RFC will receive design feedback from the larger community, and the author should be prepared to revise it in response
5. Build consensus and integrate feedback. RFCs that have broad support are much more likely to make progress than those that don't receive any comments
6. Eventually, the maintainers will decide whether the RFC is a candidate for inclusion in VBA-Web
7. RFCs that are candidates for inclusion in VBA-Web will enter a "final comment period" lasting 7 days. The beginning of this period will be signaled with a comment and tag on the RFC's pull request
8. An RFC can be modified based upon feedback from the maintainers and community. Significant modifications may trigger a new final comment period.
9. An RFC may be rejected by the maintainers after public discussion has settled and comments have been made summarizing the rationale for rejection. A member of the maintainers should then close the RFC's associated pull request.
10. An RFC may be accepted at the close of its final comment period. A maintainer will merge the RFC's associated pull request, at which point the RFC will become 'active'.

## The RFC Lifecycle

Once an RFC becomes active then authors may implement it and submit the feature as a pull request to the VBA-Web repo. Becoming 'active' is not a rubber stamp, and in particular still does not mean the feature will ultimately be merged; it does mean that the core team has agreed to it in principle and are amenable to merging it.

Furthermore, the fact that a given RFC has been accepted and is 'active' implies nothing about what priority is assigned to its implementation, nor whether anybody is currently working on it.

Modifications to active RFC's can be done in followup PR's. We strive to write each RFC in a manner that it will reflect the final design of the feature; but the nature of the process means that we cannot expect every merged RFC to actually reflect what the end result will be at the time of the next major release; therefore we try to keep each RFC document somewhat in sync with the language feature as planned, tracking such changes via followup pull requests to the document.