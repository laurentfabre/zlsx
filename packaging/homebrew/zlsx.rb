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
  # Keep version synced with build.zig.zon — the smoke `test do` block
  # exercises both `zlsx` and `zlsx-extract-images`, and the latter
  # only ships in 0.3.0+ tarballs. If you cherry-pick this template
  # into a tap pinned to an older release, also drop the
  # zlsx-extract-images line from the test block.
  version "0.4.0"
  # Proprietary (commercial + 60-day eval). Distribute via personal tap
  # only; homebrew/core requires OSI-approved licenses.
  license :cannot_represent

  on_macos do
    if Hardware::CPU.arm?
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-aarch64-apple-darwin.tar.gz"
      sha256 "174fb88053ec3a6fce17cbf4541cd05c99578483b9ee8c5856613e925ed6cbe7"
    else
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-x86_64-apple-darwin.tar.gz"
      sha256 "1564d72d0e7aab99c73de5c8985dac7baa2632f1f3f56a7a6270e87ce8c7d548"
    end
  end

  on_linux do
    if Hardware::CPU.arm?
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-aarch64-linux-musl.tar.gz"
      sha256 "9131587397260d6bb252fbe550955bf85e500af9710689985ab34e261aecddbc"
    else
      url "https://github.com/laurentfabre/zlsx/releases/download/v#{version}/zlsx-#{version}-x86_64-linux-musl.tar.gz"
      sha256 "66e8f4599943bfd8a0e3c1344d72b04b5aa7fea9b551e494fa5b6af39445b3b1"
    end
  end

  def install
    # 0.3.0+ tarballs ship two binaries: the main `zlsx` CLI and
    # `zlsx-extract-images` (standalone OOXML image extractor that
    # uses the new zlsx_pkg package layer). `bin.install Dir["bin/*"]`
    # picks up both regardless of platform suffix.
    bin.install Dir["bin/*"]
    lib.install Dir["lib/*"]
    include.install "include/zlsx.h"
    doc.install "README.md"
  end

  test do
    # Basic sanity: --help prints usage.
    assert_match "usage: zlsx", shell_output("#{bin}/zlsx --help")
    # extract-images binary is present (0.3.0+) and prints its usage
    # banner on stderr when invoked with no args. Exit code 1 is the
    # documented bad-usage signal.
    assert_match "usage:", shell_output("#{bin}/zlsx-extract-images 2>&1", 1)
  end
end
