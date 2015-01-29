#include <QtTest>
#include <QString>
#include <QDebug>
#include <QByteArray>
#include <QFile>
#include <QFont>
#include <QTime>
#include <QColor>
#include <QDomDocument>

#include <docx/document.h>
#include <docx/text.h>


using namespace Docx;

class TestDocument : public QObject
{
    Q_OBJECT

public:
    TestDocument();

private Q_SLOTS:
    void testLoad();
};


TestDocument::TestDocument()
{

}

void TestDocument::testLoad()
{
    //demo.docx  Empty.docx
    Document doc(QStringLiteral("Empty.docx"));
    Paragraph *p = doc.addParagraph("helleWord", "");
    p->insertParagraphBefore("Before", "");
    Run *addRun = p->addRun("AddRun");
    addRun->addText("Hello");
    doc.save("aSave.docx");
}


QTEST_APPLESS_MAIN(TestDocument)
#include "main.moc"
