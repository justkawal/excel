part of excel;

Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> _archiveFiles, {
  String? excludedFile,
}) {
  var clone = Archive();
  
  // First copy existing files
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
  
  // Then add any new files from _archiveFiles that weren't in the original archive
  _archiveFiles.forEach((name, file) {
    if (!archive.files.any((f) => f.name == name)) {
      var compress = !_noCompression.contains(name);
      file.compress = compress;
      clone.addFile(file);
    }
  });
  
  return clone;
}
