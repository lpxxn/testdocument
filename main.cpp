#include <QtTest>
#include <QString>
#include <QDebug>
#include <QByteArray>
#include <QFile>
#include <QFont>
#include <QTime>
#include <QColor>
#include <QDomDocument>

#include "docx/document.h"


using namespace Docx;

class TestDocument : public QObject
{
    Q_OBJECT

public:
    TestDocument();

public Q_SLOTS:
    void testLoad();
};


TestDocument::TestDocument()
{

}

void TestDocument::testLoad()
{
    Document doc(QStringLiteral("aaa.docx"));

}


QTEST_APPLESS_MAIN(TestDocument)
#include "main.moc"
