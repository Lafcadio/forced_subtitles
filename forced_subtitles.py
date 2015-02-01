from gdata.spreadsheet.service import SpreadsheetsService
import pickle
import time
import re, string

DEBUG = False


class SheetQuerier:
    key = "0AkGO8UqErL6idDhYYjg1ZXlORnRaM3ZhTks4Z3FrYlE"
    worksheets = {"DVD" : "title", "Blu-ray" : "title"}
    cache = "spreadsheet_cache.p"
    cache_timeout = 60*60*24*7
    prefixes = ["The", "A"]
    whitespace_pattern = re.compile('[^a-zA-Z0-9 ]+', re.UNICODE)

    def __init__(self, force_reload = False):
        try:
            pickle_date, self.data = pickle.load( open( self.cache, "rb" ) )
            if DEBUG:
                print "Loaded cache successfully"
            if force_reload or time.time() - pickle_date > self.cache_timeout:
                if DEBUG:
                    if force_reload:
                        print "Cache update forced."
                    else:
                        print "Cache too old, reloading."
                self.reload_data()
                pickle.dump( (time.time(), self.data), open( self.cache, "wb"))
        except:
            if DEBUG:
                print "Problem loading cache, reloading."
            self.reload_data()
            pickle.dump( (time.time(), self.data), open( self.cache, "wb"))


    def reload_data(self):
        self.client = SpreadsheetsService()
        feed = self.client.GetWorksheetsFeed(self.key, visibility='public', projection='basic')
        self.data = {}
        for entry in feed.entry:
            if entry.title.text in self.worksheets.keys():
                bad_rows,total_rows = self.process_sheet(entry.id.text.split("/")[-1], self.worksheets[entry.title.text])
                print "Skipped %d / %d rows in sheet \"%s\"" % (bad_rows, total_rows, entry.title.text)
            elif DEBUG:
                print "Skipped sheet \"%s\"" % entry.title.text


    def process_sheet(self, sheet_key, movie_row_key, type_row_key = "forcedsubtitletype"):
        if DEBUG:
            print "Document: %s" % self.key
            print "Sheet: %s" % sheet_key

        rows = self.client.GetListFeed(self.key, sheet_key, visibility='public', projection='values').entry
        bad_rows = 0
        for row in rows:
            try:
                self.data[SheetQuerier.clean_title(row.custom[movie_row_key].text.strip())] = row.custom[type_row_key].text.strip()
            except:
                bad_rows += 1
        return bad_rows, len(rows)

    def query_exact(self, title):
        query = SheetQuerier.clean_title(title)
        if query in self.data:
            return self.data[query]
        else:
            return False

    @staticmethod
    def clean_title(title):
        # Move prefixes
        for prefix in SheetQuerier.prefixes:
            if title.endswith(prefix):
                title = "%s %s" % (prefix, title[:-1 * len(prefix) - 2])
                break
        # Strip all non alpha-numeric characters
        title = SheetQuerier.whitespace_pattern.sub('', title)

        # Return the lowercase version
        return title.lower()


q = SheetQuerier()
