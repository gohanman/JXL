#/bin/sh

mvn exec:java -Dexec.mainClass="coop.wholefoods.jxl.App" -Dexec.classpathScope=runtime -Dexec.args="$@"

