import mysql.connector as mysql

ERRORS = True
PROGRESS = True


class connectDatabase:
    def __init__(self, host, err=ERRORS, prog=PROGRESS):
        if host == "cloud":
            self.dbHost = "192.168.30.101"
            self.dbUser = "remote"
            self.dbPass = "Framing123.password"
            self.dbDb = "ja_db"

        elif host == "local":
            self.dbHost = "localhost"
            self.dbUser = "root"
            self.dbPass = ""
            self.dbDb = "test_db"
        print(
            "Connecting to database...\n"
            "\tHost: %s\n"
            "\tUser: %s\n"
            "\tDatabase: %s\n"
            "\t________________________________" % (self.dbHost, self.dbUser, self.dbDb)
        )

        self.dbConn = mysql.connect(
            host=self.dbHost,
            user=self.dbUser,
            passwd=self.dbPass,
            db=self.dbDb,
            use_pure=True,
        )
        self.dbCursor = self.dbConn.cursor()

    def reconnect(self):
        self.dbConn = mysql.connect(
            host=self.dbHost,
            user=self.dbUser,
            passwd=self.dbPass,
            db=self.dbDb,
            use_pure=True,
        )
        self.dbCursor = self.dbConn.cursor()
        print("Reconnected to database...\n")

    def connectionTimeout(self, timeout=6000):

        self.dbCursor.execute("SET GLOBAL connect_timeout=%s" % timeout)
        self.dbConn.commit()
