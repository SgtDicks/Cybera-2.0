# Third-party and source notices

The modern bridge publish output is built from this project and the Microsoft .NET 8 runtime supplied by the target server's .NET installation. The installer is compiled with Inno Setup 6, which must be installed separately on the build machine; see its license and notices at <https://jrsoftware.org/isinfo.php>.

The installer intentionally does not bundle Cybera 1.9.rc2 or deploy it to clients. The untouched original Cybera source remains in `third_party/original-source/`; its GNU GPL version 2 notice is preserved in `license.txt`. If a future distribution bundles Cybera or modified Cybera source, include the applicable GPL notice and corresponding source according to its terms.
