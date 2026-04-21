# Homebrew formula for the zlsx CLI.
#
# This file is a template meant to live in a personal Homebrew tap —
# typically https://github.com/<user>/homebrew-zlsx — so users install
# it with:
#
#   brew tap <user>/zlsx
#   brew install zlsx
#
# On each release tag, the release workflow publishes per-platform
# tarballs to GitHub Releases. Bump `version` below and update the two
# macOS sha256 entries (arm and intel) from the SHA256SUMS asset, then
# commit to the tap repo.
#
# The formula ships only the CLI binary. The C library (libzlsx.dylib,
# libzlsx.a) and the header (include/zlsx.h) are also included in the
# tarball so downstream C consumers can install them manually.

class Zlsx < Formula
  desc "Tiny, read-only .xlsx parser + CLI (Zig, no third-party deps)"
  homepage "https://github.com/laurentfabre/zlsx"
  version "0.2.0"
  license "MIT"

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-aarch64-apple-darwin.tar.gz"
      sha256 "9cb20ed73b0217f5d9d4b90a4234507ec6d7d54fbd778390cb8630ca1d874cf0"
    else
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-x86_64-apple-darwin.tar.gz"
      sha256 "2ab5b27a4e019d4ec6c9d3f716c3b7935da282cb153fcaafefbb424bfd045966"
    end
  end

  on_linux do
    if Hardware::CPU.arm?
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-aarch64-linux-musl.tar.gz"
      sha256 "b3a85a23159add32ed916d43395bd3e44cf6d5393e0796352376ff5c83b875ce"
    else
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-x86_64-linux-musl.tar.gz"
      sha256 "698a9c223fe36246c6d7e7ed05b44ae2e27ebc1ca7b0f9cd7c2f6e702c47be30"
    end
  end

  def install
    bin.install "bin/zlsx"
    lib.install Dir["lib/*"]
    include.install "include/zlsx.h"
    doc.install "README.md"
  end

  test do
    # Basic sanity: --help prints usage.
    assert_match "usage: zlsx", shell_output("#{bin}/zlsx --help")
  end
end
