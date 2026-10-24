"""Exception types for kitchensink4ppt. Every tool call maps these to actionable messages."""

from __future__ import annotations

from typing import Any


CODES = frozenset({
    "AMBIGUOUS_LOCATION",
    "APP_BLOCKED",
    "APP_BUSY",
    "APP_NOT_RUNNING",
    "BAD_ARGS",
    "BAD_PARAMS",
    "CONFLICT",
    "DOCUMENT_LOCKED",
    "INTERNAL_ERROR",
    "LIMIT_EXCEEDED",
    "NOT_AVAILABLE",
    "NOT_FOUND",
    "PROTECTED",
    "PROTECTED_VIEW",
    "STALE_ANCHOR",
    "TIMEOUT",
    "UNSUPPORTED_CONTENT",
    "VALIDATION_FAILED",
})


class OpError(Exception):
    """A surface refusal with one closed code and bounded structured detail."""

    def __init__(
        self,
        code: str,
        message: str,
        *,
        hint: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code if code in CODES else "INTERNAL_ERROR"
        self.message = message
        self.hint = hint
        self.details = details


class PptMcpError(Exception):
    """Base class; message text is user-facing."""


class DocumentNotFound(PptMcpError):
    pass


class DocumentLocked(PptMcpError):
    """File is open in PowerPoint (or another process holds a lock)."""


class DocumentHeld(DocumentLocked):
    """Another program holds the deck, or the new copy of it, so a save's
    swap was refused and nothing changed. PowerPoint does not hold it: its
    owner file is checked first and is not there. So a live='auto' call does
    not go looking for the deck in PowerPoint; the caller retries."""


class DeckSetAside(PptMcpError):
    """A save could not finish and could not put the deck back in its place.
    The message says where the deck is; its content is safe there, and the
    next change or the next start of the server puts it back."""

    code = "CONFLICT"


class DocumentBusy(PptMcpError):
    """Another program was writing the file while it was being read (a save
    replacing it). Not a DocumentLocked, so a live='auto' call does not go
    looking for the deck in PowerPoint; the caller retries instead."""

    hint = "Retry the call; the other program's save finishes in a moment."


class DocumentReadOnly(PptMcpError):
    """The file's read-only attribute prevents an in-place mutation."""


class DocumentCorrupt(PptMcpError):
    """File is not a valid .pptx package."""


class DocumentProtected(PptMcpError):
    """File is encrypted/password-protected."""


class DocumentEditProtected(PptMcpError):
    """Presentation carries a p:modifyVerifier edit password."""


class TargetNotFound(PptMcpError):
    """Slide, shape, placeholder, or anchor the tool was told to act on does not exist."""


class AmbiguousTarget(PptMcpError):
    """More than one match for the given anchor; caller must disambiguate."""


class UnsupportedStructure(PptMcpError):
    """Package topology or XML shape we refuse to guess about (conservative mode)."""


class UnsupportedPackage(UnsupportedStructure):
    """A package laid out in a way this server does not read: saved as Strict
    Open XML, or with its presentation under another part name. Refused when
    the deck is opened, by every read and every edit, instead of being read
    as a deck with no slides or from the wrong part."""

    hint = (
        "Open the deck in PowerPoint and save it as a PowerPoint Presentation (.pptx), "
        "then use the new file."
    )


class ValidationFailed(PptMcpError):
    """Post-edit validation caught a problem; the original file was NOT modified."""


class SlideIdsExhausted(PptMcpError):
    """The deck's slide ids reached the highest the file format allows, so a
    new slide cannot be numbered. The deck's state stopped the call, not its
    arguments, so the hint names the step that works."""

    hint = (
        "Nothing in the request needs to change: copy this deck's slides into a new deck "
        "with merge_decks, where they are numbered afresh, and work on the new deck."
    )


class RefusedUnread(PptMcpError):
    """A file this server refuses to read for safety: a package built to
    exhaust memory, one that names a part twice (core/zipsafe.py), or a part
    that declares a document type (core/xmlsafe.py). validate never hands
    such a file to PowerPoint. It is raised as one of the two classes
    below, which keep the error codes of their first base."""


class PackageRefused(DocumentCorrupt, RefusedUnread):
    """RefusedUnread when a package is opened or its directory checked."""


class OutputRefused(ValidationFailed, RefusedUnread):
    """RefusedUnread when a payload is validated (a save's output, a backup,
    a restore source, the file given to validate)."""


class PowerPointNotRunning(PptMcpError):
    """No attachable interactive PowerPoint instance (live tools need one)."""


class DocumentNotOpenInPowerPoint(PptMcpError):
    """Live tool targeted a presentation that is not open in the running PowerPoint."""


class ProtectedViewRefused(PptMcpError):
    """Presentation is in Protected View; the user must click Enable Editing."""


class PowerPointBusy(PptMcpError):
    """PowerPoint rejected the call (dialog, Backstage, or a running command)."""


class PowerPointBlocked(PptMcpError):
    """PowerPoint is not answering at all (long synchronous operation in progress)."""


class PowerPointDisconnected(PptMcpError):
    """PowerPoint or the presentation closed mid-call; the edit may be partially applied."""


class LiveLockTimeout(PptMcpError):
    """Another server process held the cross-process live-session lock past
    the wait window; the live call was refused and nothing was changed."""
