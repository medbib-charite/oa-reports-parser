# oa-reports-parser

Publikationsdaten über Veröffentlichungen in bestehenden Verlagsverträgen werden in sehr unterschiedlichen Formaten von Verlagen an Bibliotheken gemeldet. Die Daten werden als Excel- oder seltener CSV-Dateien bereitgestellt, dabei verwendet jeder Verlag ein eigenes Schema. An der Medizinischen Bibliothek der Charité – Universitätsmedizin Berlin wurden zwei Skripte geschrieben, die von einer einfachen Konfigurationsdatei ausgehend aus Dashboard- und sonstigen Reports der Verlage Excel-Dateien in das von der Charité gewünschte Schema übertragen.  

Um das Skript besser testen zu können, stellen wir neben dem Sourcecode auch sog. [Mockdata](https://github.com/medbib-charite/oa-reports-parser/tree/main/mock_data) zur Verfügung.

## dare2puli.py: dashboard report to publisher list

Entwickelt wurde das Skript, um die Dashboard-Reports der drei DEAL-Verlage (Elsevier, Springer Nature, Wiley) in das von der Charité gewünschte Format zum Publikationsmonitoring zu bringen. Das Skript überträgt die Inhalte einzelner Excel-Dateien, die man als Report aus den DEAL-Dashboards ziehen kann, in das gewünschte Schema zum Publikationsmonitoring.

Das gewünschte Zielschema wird in der Datei `config/mock_mapping.csv` festgelegt. Spalten, die keine Entsprechung in dem Dashboard-Report haben, können leer gelassen werden. Felder, die mit "#" beginnen, werden ignoriert. Einträge, die mit "-->" beginnen, werden als Default-Werte übernommen. In die Spalte _PuLi_ wird das Zielformat geschrieben; die anderen Spalten sind nach den Verlagen benannt und werden ebenfalls genutzt, um im Skript den Verlag zu identifizieren (dabei werden aus dem Dateinamen "-" in Leerzeichen umgewandelt). 

## dare2pmt.py: dashboard reports to publication management table

Das Skript wandelt Reports von bestehenden Verlagsverträgen an der Charité in das gewünschte Format zum Publikationsmonitoring um. Das Skript nimmt alle Excel- und CSV-Dateien aus einem Ordner, der per Menü ausgewählt werden kann, wandelt sie in das gewünschte Schema um und schreibt sie in eine gemeinsame Excel-Datei.

Das Skript wird an der Medizinischen Bibliothek für folgende Verlagsreports genutzt: American Chemical Society, British Medical Journals, Cambridge University Press, Hogrefe Psyjournals, Institute of Physics Publishing, Rockefeller University Press, SAGE sowie Taylor & Francis.

Um das Skript zu verwenden, muss die Konfigurationsdatei `config/mock_mapping.csv` an die individuellen Anforderungen angepasst werden, wie im Abschnitt zu dare2pmt.py beschrieben.

## missing_pub_dates.py 

Das Skript sucht fehlende Online Publication Dates mit der [Crossref-REST-API](https://www.crossref.org/documentation/retrieve-metadata/rest-api/). Dafür wird eine Excel-Datei ausgewählt. Von dieser wird im ersten Tabellenblatt die erste Tabelle als `pandas DataFrame` eingelesen. Wenn in dem Tabellenblatt keine Spalte _Online Publication Date_ existiert, wird eine angelegt. Dann wird für alle Zeilen, in denen _Online Publication Date_ keinen Eintrag enthält, die Crossref-REST-API nach Metadaten gefragt und aus ihr das Feld _published-online_ als Online Publication Date genommen. 

Es werden zwei Ergebnisdateien geschrieben: einmal wird der DataFrame als Excel-Tabelle geschrieben, einmal wird nur die Spalte Online Publiction Date als Text-Datei ausgegeben. Der Ordner, in dem sie geschrieben werden, und die Dateinamen können in dem Skript geändert werden (`RESULTS_DIR`, `OUTPUT_XLSX`, `OUTPUT_TXT`) - für ein solches kleines Skript erschien eine Konfigurationsdatei zu umfänglich. 

## Fragen

Fragen oder Feedback: openaccess@charite.de. 

## Lizenz

Die von Sebastian Dittmann ([github.com/sebdit](https://github.com/sebdit)) entwickelten Skripte steht unter [MIT-Lizenz](https://github.com/medbib-charite/oa-reports-parser/blob/main/LICENSE).
