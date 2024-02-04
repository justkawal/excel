part of excel;

Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> _archiveFiles, {
  String? excludedFile,
}) {
  var clone = Archive();
  archive.files.forEach((file) {
    if (file.isFile) {
      if (excludedFile != null &&
          file.name.toLowerCase() == excludedFile.toLowerCase()) {
        return;
      }
      ArchiveFile copy;
      if (_archiveFiles.containsKey(file.name)) {
        copy = _archiveFiles[file.name]!;
      } else {
        var content = file.content as Uint8List;
        var compress = !_noCompression.contains(file.name);
        copy = ArchiveFile(file.name, content.length, content)
          ..compress = compress;
      }
      clone.addFile(copy);
    }
  });
  return clone;
}
