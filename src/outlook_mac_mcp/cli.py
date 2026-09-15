import argparse
import sys
from collections.abc import Sequence

from outlook_mac_mcp.bootstrap import build_use_cases
from outlook_mac_mcp.domain.errors import OutlookMcpError
from outlook_mac_mcp.infrastructure.graph.authentication import (
    DeviceCodeAuthenticator,
    DeviceCodePrompt,
)
from outlook_mac_mcp.interface.mcp.observability import configure_logging
from outlook_mac_mcp.interface.mcp.server import build_server

SERVE_COMMAND = "serve"
SIGN_IN_COMMAND = "sign-in"
EXIT_SUCCESS = 0
EXIT_FAILURE = 1


def main(argv: Sequence[str] | None = None) -> int:
    command = _parse_command(argv if argv is not None else sys.argv[1:])
    try:
        return _run(command)
    except OutlookMcpError as error:
        print(f"{type(error).__name__}: {error}", file=sys.stderr)
        return EXIT_FAILURE


def _parse_command(argv: Sequence[str]) -> str:
    parser = argparse.ArgumentParser(prog="outlook-mac-mcp")
    parser.add_argument(
        "command",
        nargs="?",
        default=SERVE_COMMAND,
        choices=[SERVE_COMMAND, SIGN_IN_COMMAND],
        help="serve runs the MCP server on stdio; sign-in stores a token in the Keychain",
    )
    return str(parser.parse_args(list(argv)).command)


def _run(command: str) -> int:
    if command == SIGN_IN_COMMAND:
        return _sign_in()
    return _serve()


def _serve() -> int:
    """Never starts a sign-in: without a stored token the tool fails with a clear message.

    The device-code flow blocks for minutes and has nowhere to show its code here, because
    stdout carries the MCP protocol.
    """
    configure_logging()
    use_cases = build_use_cases()
    build_server(use_cases.list_unread_emails, use_cases.get_email).run()
    return EXIT_SUCCESS


def _sign_in() -> int:
    DeviceCodeAuthenticator.from_environment().sign_in(_show_prompt)
    print("Signed in. The token is stored in the macOS Keychain.", file=sys.stderr)
    return EXIT_SUCCESS


def _show_prompt(prompt: DeviceCodePrompt) -> None:
    """stderr, never stdout: the same entry point also speaks MCP over stdout."""
    minutes = prompt.expires_in_seconds // 60
    print(
        f"To sign in, open {prompt.verification_uri} and enter the code "
        f"{prompt.user_code} (valid for {minutes} minutes).",
        file=sys.stderr,
    )
