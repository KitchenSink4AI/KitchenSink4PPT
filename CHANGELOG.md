# Changelog

### 1.2.1
- The server moves to the same foundation library as its siblings (fastmcp 3), so all four KitchenSink servers now install together into one environment. Tiered loading is unchanged and re-verified over the wire: packs still switch mid-session and clients are still told the tool list changed.

### 1.2.0
- Every read is budgeted: one call can no longer flood a conversation, and anything left out is named with a continuation.
- Saves are up to three times faster and now preserve each media file's original compression exactly.
- Update notice: a weekly, disclosed check for newer releases, off with KS4P_UPDATE_CHECK=off.
- The install screen and info card are rewritten in plain language.
