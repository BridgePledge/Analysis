import requests
import bs4
import xlsxwriter


import collections
import csv
import json
import os
import sqlite3
import time

######################################################################
## Helpers

# h/t https://stackoverflow.com/a/58106295/986793
def with_attrs(**kwargs):
    def decorator(func):
        for k, v in kwargs.items():
            setattr(func, k, v)
        return func
    return decorator

def make_fetcher(url, output):

    @with_attrs(url=url, output=output)
    def fetcher():
        if os.path.isfile(fetcher.output):
            return
        print("fetching: ", fetcher.url)
        r = requests.get(fetcher.url, headers={'User-Agent': 'Mozilla/5.0'})
        with open(fetcher.output, "w") as out:
            out.write(r.text)
        time.sleep(2)

    return fetcher


######################################################################

scrape_538_quiet_caucus = make_fetcher(
        url= "https://github.com/fivethirtyeight/data/raw/master/quiet-caucus-2024/clustered_congress.csv",
        output= "data/raw/clustered_congress.csv"
)

def write_538(con, row):
    con.execute("""
    insert into five_thirty_eight_quiet_caucus (
        chamber, icpsr, name, last_name , party , district , terms_pct , margin_2020 , cluster , agreement_score , progressive , new_dems , blue_dogs , problem_solvers , RMSP , governance , study , freedom , dw_nominate_dim1 , dw_nominate_dim2 , notes
    ) values (
        'House', ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?,  ?, ?
    );
    """, row )

def load_538(con, filename):
    with con:
        con.execute("""

        CREATE TABLE five_thirty_eight_quiet_caucus (
            chamber text,
            icpsr integer,
			name text,
			last_name text,
			party text,
			district text,
			terms_pct real,
			margin_2020 text,
			cluster text,
			agreement_score real,
			progressive integer,
			new_dems integer,
			blue_dogs integer,
			problem_solvers integer,
			RMSP integer,
			governance integer,
			study integer,
			freedom integer,
			dw_nominate_dim1 real,
			dw_nominate_dim2 real,
			notes text
        )
        """)

        with open(filename) as fp:
            reader = csv.reader(fp)
            next(reader) # throw out header
            for row in reader:
                row[0] = int(row[0])
                row[5] = float(row[5].replace("%", "")) if row[5] != "ƒis" else None #Absolutely bizzare data for Jim Himes
                row[8] = float(row[8].replace("%", "")) if row[8] != "" else None
                for i in range(9,17):
                    row[i] = 0 if row[i] == "" else 1
                row[17] = float(row[17])
                row[18] = float(row[18])
                write_538(con, row)

######################################################################


scrape_cook_pvi = make_fetcher(
        url= "https://datawrapper.dwcdn.net/IERCx/2/dataset.csv",
        output= "data/raw/cook_pvi.csv"
)

def write_cook_pvi(con, row):
        con.execute("""
        insert into cook_pvi (
            Dist, First, Last, Party, PVI_2023, PVI_2023_raw, Rank
        ) values (
             ?,  ?,  ?,  ?,  ?,  ?, ?
        );
        """, row)


def load_cook_pvi(con, filename):
    with con:
        con.execute("""
        CREATE TABLE cook_pvi (
            Dist text,
            First text,
            Last text,
            Party text,
            PVI_2023 text,
            PVI_2023_raw text,
            Rank integer
            );
        """)
        with open(filename) as fp:
            reader = csv.reader(fp, delimiter='\t')
            next(reader) # throw out header
            for row in reader:
                row[0] = row[0][:5].replace("-0", "-")
                row[2] = row[2].replace("í", "i").replace("á", "a")
                row[6] = int(row[6])
                write_cook_pvi(con, row)
######################################################################

scrape_govtrack_cosponsor_house = make_fetcher(
        url= "https://www.govtrack.us/congress/members/report-cards/2022/house/cosponsored-other-party.csv",
        output= "data/raw/govtrack_cosponsor_house.csv"
)

scrape_govtrack_cosponsor_senate = make_fetcher(
        url= "https://www.govtrack.us/congress/members/report-cards/2022/senate/cosponsored-other-party.csv",
        output= "data/raw/govtrack_cosponsor_senate.csv"
)

def write_govtrack_cosponsor(con, row, chamber):
        con.execute("""
        insert into govtrack_cosponsor (
            chamber, rank_from_low, rank_from_high, percentile, cosponsored_other_party, id, bioguide_id, state, district, name
        ) values (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        );
        """, (chamber, *row))

def load_govtrack_cosponsor(con, filename, chamber):
    with con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS govtrack_cosponsor(
            chamber text,
            rank_from_low int,
			rank_from_high int,
			percentile int,
			cosponsored_other_party real,
			id int,
			bioguide_id text,
			state text,
			district int,
			name text
            );
        """)
        with open(filename) as fp:
            reader = csv.reader(fp)
            next(reader) # throw out header
            for row in reader:
                row[0] = int(row[0])
                row[1] = int(row[1])
                row[2] = int(row[2])
                row[3] = float(row[3])
                row[4] = int(row[4])
                row[7] = int(row[7]) if row[7] != "" else None
                row[8] = eval(row[8]).decode("utf-8").replace("í","i").replace("á", "a") # convert bytes escaped to regular text
                write_govtrack_cosponsor(con, row, chamber)
######################################################################

scrape_voteview = make_fetcher(
        url= "https://voteview.com/static/data/out/members/HSall_members.csv",
        output= "data/raw/vote_view.csv"
)

def write_voteview(con, row):
        con.execute("""
        insert into voteview (
            congress, chamber, icpsr, state_icpsr, district_code, state_abbrev, party_code, occupancy, last_means, bioname , bioguide_id, born, died, nominate_dim1, nominate_dim2, nominate_log_likelihood, nominate_geo_mean_probability, nominate_number_of_votes, nominate_number_of_errors, conditional, nokken_poole_dim1, nokken_poole_dim2
        ) values (
           ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
        );
        """, row)


def load_voteview(con, filename):
    with con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS voteview(
            congress int,
			chamber text,
			icpsr int,
			state_icpsr int,
			district_code int,
			state_abbrev text,
			party_code int,
			occupancy int,
			last_means int,
			bioname  text,
			bioguide_id text,
			born int,
			died int,
			nominate_dim1 real,
			nominate_dim2 real,
			nominate_log_likelihood real,
			nominate_geo_mean_probability real,
			nominate_number_of_votes int,
			nominate_number_of_errors int,
			conditional text,
			nokken_poole_dim1 real,
            nokken_poole_dim2 real
            );
        """)
        with open(filename) as fp:
            reader = csv.reader(fp)
            next(reader) # throw out header
            for row in reader:
                row[0] = int(row[0])

                if row[0] < 118: continue # only want most recent congress

                row[2] = int(row[2])
                row[3] = int(row[3])
                row[4] = int(row[4])
                row[6] = int(row[6]) # party code
                row[7] = int(row[7]) if row[7] != '' else None # occupancy
                row[8] = int(row[8]) if row[8] != '' else None # last_means
                row[9] = row[9].replace("Á", "A").replace("Ó", "O").replace("Í","I")
                row[11] = int(row[11][:4]) if row[11] != '' else None # born
                row[12] = int(row[12][:4]) if row[12] != '' else None # died
                for i in range(13, 17):
                    row[i] = float(row[i]) if row[i] else None
                row[17] = int(row[17]) if row[17] else None # num_votes
                row[18] = int(row[18]) if row[18] else None # of_errors
                row[20] = float(row[20])
                row[21] = float(row[21])
                write_voteview(con,row)
######################################################################

scrape_lugar_house = make_fetcher(
        url= "https://www.thelugarcenter.org/ourwork-86.html",
        output= "data/raw/lugar_house.html"
)

scrape_lugar_senate = make_fetcher(
        url= "https://www.thelugarcenter.org/ourwork-85.html",
        output= "data/raw/lugar_senate.html"
)

def write_lugar(con, row, chamber):
        con.execute("""
        insert into lugar (
            chamber, first, last, state, party, score
        ) values (
           ?,?,?,?,?,?
        );
        """, (chamber, *row))


def load_lugar(con, filename, chamber):
    with con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS lugar(
            chamber text,
            first text,
            last text,
            state text,
            party text,
            score real
            );
        """)

        with open(filename) as fp:
            soup = bs4.BeautifulSoup(fp)

        main_table = soup.find_all("table")[2]

        rows = iter(main_table.find_all("tr"))
        next(rows) # skip top header
        next(rows) # skip inner header

        for row in rows:
            row = [x.text for x in row.find_all("td")][1:6]
            row[4] = float(row[4])
            write_lugar(con, row, chamber)

######################################################################

states = [
    ('ALABAMA', 'AL'),
    ('ALASKA', 'AK'),
    ('AMERICAN SAMOA', 'AS'),
    ('ARIZONA', 'AZ'),
    ('ARKANSAS', 'AR'),
    ('CALIFORNIA', 'CA'),
    ('COLORADO', 'CO'),
    ('CONNECTICUT', 'CT'),
    ('DELAWARE', 'DE'),
    ('DISTRICT OF COLUMBIA', 'DC'),
    ('FLORIDA', 'FL'),
    ('GEORGIA', 'GA'),
    ('GUAM', 'GU'),
    ('HAWAII', 'HI'),
    ('IDAHO', 'ID'),
    ('ILLINOIS', 'IL'),
    ('INDIANA', 'IN'),
    ('IOWA', 'IA'),
    ('KANSAS', 'KS'),
    ('KENTUCKY', 'KY'),
    ('LOUISIANA', 'LA'),
    ('MAINE', 'ME'),
    ('MARYLAND', 'MD'),
    ('MASSACHUSETTS', 'MA'),
    ('MICHIGAN', 'MI'),
    ('MINNESOTA', 'MN'),
    ('MISSISSIPPI', 'MS'),
    ('MISSOURI', 'MO'),
    ('MONTANA', 'MT'),
    ('NEBRASKA', 'NE'),
    ('NEVADA', 'NV'),
    ('NEW HAMPSHIRE', 'NH'),
    ('NEW JERSEY', 'NJ'),
    ('NEW MEXICO', 'NM'),
    ('NEW YORK', 'NY'),
    ('NORTH CAROLINA', 'NC'),
    ('NORTH DAKOTA', 'ND'),
    ('NORTHERN MARIANA IS', 'MP'),
    ('OHIO', 'OH'),
    ('OKLAHOMA', 'OK'),
    ('OREGON', 'OR'),
    ('PENNSYLVANIA', 'PA'),
    ('PUERTO RICO', 'PR'),
    ('RHODE ISLAND', 'RI'),
    ('SOUTH CAROLINA', 'SC'),
    ('SOUTH DAKOTA', 'SD'),
    ('TENNESSEE', 'TN'),
    ('TEXAS', 'TX'),
    ('UTAH', 'UT'),
    ('VERMONT', 'VT'),
    ('VIRGINIA', 'VA'),
    ('VIRGIN ISLANDS', 'VI'),
    ('WASHINGTON', 'WA'),
    ('WEST VIRGINIA', 'WV'),
    ('WISCONSIN', 'WI'),
    ('WYOMING', 'WY')
]

def commonground_is_congress(politician):
    return None is not (
       politician.find(class_="U.S. House")
       or politician.find(class_="U.S. Senate")
    )

@with_attrs(base="data/raw/commonground")
def scrape_commonground_state_index(state_tup):
    state, abbr = state_tup
    url = f"https://commongroundscorecard.org/tag/{abbr}/"
    out_dir = f"{scrape_commonground_state_index.base}/{abbr}/"
    out_index = f"{out_dir}/index.html"
    os.makedirs(out_dir, exist_ok=True)

    make_fetcher(url=url, output=out_index)()

    with open(out_index) as f:
        soup = bs4.BeautifulSoup(f.read())

    for politician in soup.find_all(class_="politician-wrap"):
        if not commonground_is_congress(politician):
            continue
        url = politician.find('a')['href']
        stub = url.split('/')[-2]
        output = f"{out_dir}/{stub}.html"
        make_fetcher(url=url, output=output)()


def load_commonground(con, filename):
    with con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS commonground(
            name text,
            chamber text,
            state text,
            district int,
            party text,
            total int,
            official_performance int,
            personal_actions int,
            communications int,
            commitments int,
            bonus int,
            constituent_partisan_intensity int
            );
        """)

    with con:
        for d, _, files in os.walk(filename):
            for f in files:
                if f == "index.html":
                    continue
                row = read_commonground_scorecard(f"{d}/{f}")
                write_commonground(con, row)

def read_commonground_scorecard(filename):
    with open(filename) as fp:
        soup = bs4.BeautifulSoup(fp)

    h = soup.head.title.text
    name =  h[:h.find(" - ")]

    h = soup.article.h5.text
    chamber = h[h.find(" ")+1:]

    h = soup.find("span", class_="state-district-pary").text
    h = h.replace("(", "").replace(" ", "").replace(")", "").split("-")
    if chamber == "House":
        state, district, party =  h
        district = int(district)
    else:
        if "" in h and len(h) == 3:
            h.remove("")
        state, party = h
        district = None



    rightbar = soup.find("div", class_="single-main-right")

    total = int(rightbar.h5.text)

    subscores = rightbar.find_all("div", class_="single-official-performance")
    official_performance = int(subscores[0].div.div.text)
    personal_actions     = int(subscores[1].div.div.text)
    communications       = int(subscores[2].div.div.text)
    commitments          = int(subscores[3].div.div.text)
    bonus                = int(subscores[4].div.div.text)

    mavmat_script = soup.find("div", class_="mm-container").find("script", src=None).text
    left = mavmat_script.find("arrayToDataTable(") + 17
    right = mavmat_script.find(");", left)
    #mavmat = json.loads(mavmat_script[left:right])
    mavmat = eval(mavmat_script[left:right].replace("[ ,", "["))

    constituent_partisan_intensity = mavmat[1][0]

    return (
            name,
            chamber,
            state,
            district,
            party,
            total,
            official_performance,
            personal_actions,
            communications,
            commitments,
            bonus,
            constituent_partisan_intensity
            )

def write_commonground(con, row):
    con.execute("""
    insert into commonground (
        name, chamber, state, district, party, total, official_performance, personal_actions, communications, commitments, bonus, constituent_partisan_intensity
    ) values (
       ?,?,?,?,?,?,?,?,?,?,?,?
    )
    """, row)

######################################################################
######################################################################

def get_all(con):
    query = """

    with fte as (
      select *
      from five_thirty_eight_quiet_caucus
    ),
    cook_clean as (
      select concat(First, ' ', Last) as name, *
      from cook_pvi

    ),
    govtrack_clean as (
      select concat(state, "-", case when chamber = "House" then district else 0 end) as state_district,
      *
      from govtrack_cosponsor
    ),
    voteview_clean as (
      select *,
        (substr(bioname, 1, instr(bioname, ',')-1)) as last,
        concat(state_abbrev, "-", district_code) as state_district
      from voteview
      where not icpsr in (29373, 90915) -- Manchin and Menendez have extra records?
    ),
    lugar_clean as (
      select * from lugar
    ),
    commonground_clean as (
      select concat(state, "-", case when chamber = "House" then district else 0 end) as state_district,
        substr(name, instr(name, " ")+1, 999) as last,
      *
      from commonground
    )

    select * from fte
    full outer join cook_clean on (
            fte.name = cook_clean.name
            or (
              fte.last_name = cook_clean.last
              and fte.District = cook_clean.Dist
              and fte.party = cook_clean.party
            )
                       -- or (instr(fte.name, cook_clean.Last) and substr(fte.name,1,1) = substr(cook_clean.First, 1,1))
        )
    full outer join govtrack_clean on (
          (
            cook_clean.last = govtrack_clean.name
            AND govtrack_clean.state_district = cook_clean.Dist
          ) or (
            fte.last_name = govtrack_clean.name
            AND govtrack_clean.state_district = fte.District
          )
        )
    full outer join voteview_clean on (
          (
            upper(voteview_clean.last) = upper(fte.last_name)
            AND voteview_clean.state_district = fte.District
          ) or (
            upper(voteview_clean.bioname) like upper(concat(cook_clean.last,", ",cook_clean.first, "%"))
            AND voteview_clean.state_district = cook_clean.Dist
          ) or (
            voteview_clean.bioguide_id = govtrack_clean.bioguide_id
          )
        )
    full outer join lugar_clean on (
      (
        fte.last_name = lugar_clean.last
        and fte.name like Concat(substr(lugar_clean.first,1,1), "%")
        and fte.district like Concat(lugar_clean.state, "-%")
        and fte.party = lugar_clean.party
        AND fte.chamber = lugar_clean.chamber
      ) or (
        govtrack_clean.name = lugar_clean.last
        AND govtrack_clean.state = lugar_clean.state
        AND govtrack_clean.chamber = lugar_clean.chamber
      ) or (
        upper(voteview_clean.bioname) like upper(concat(lugar_clean.last, ", ", substr(lugar_clean.first,1,1), "%"))
        AND voteview_clean.state_abbrev = lugar_clean.state
      )
    )
    full outer join commonground_clean on (
      (
        voteview_clean.state_district = commonground_clean.state_district
        and upper(commonground_clean.last) like concat("%",upper(voteview_clean.last),"%")
      ) or (
        govtrack_clean.state_district = commonground_clean.state_district
        and commonground_clean.last like concat("%",govtrack_clean.name,"%")
      )
    )
    """

    with con:
        cur = con.cursor()
        result = cur.execute(query)
        title = [i[0] for i in cur.description]
        return title, result.fetchall()



######################################################################
######################################################################

if __name__ == "__main__":
    os.makedirs("data/raw", exist_ok=True)
    os.makedirs("data/clean", exist_ok=True)


    # Scrape
    scrape_538_quiet_caucus()
    scrape_cook_pvi()
    scrape_govtrack_cosponsor_house()
    scrape_govtrack_cosponsor_senate()
    scrape_voteview()
    scrape_lugar_house()
    scrape_lugar_senate()
    collections.deque(map(scrape_commonground_state_index, states), 0)


    # LOAD

    con = sqlite3.connect("data/clean/bridge.sqlite")

    load_538(con, scrape_538_quiet_caucus.output)
    load_cook_pvi(con, scrape_cook_pvi.output)
    load_govtrack_cosponsor(con, scrape_govtrack_cosponsor_house.output, "House")
    load_govtrack_cosponsor(con, scrape_govtrack_cosponsor_senate.output, "Senate")
    load_voteview(con, scrape_voteview.output)
    load_lugar(con, scrape_lugar_house.output, 'House')
    load_lugar(con, scrape_lugar_senate.output, 'Senate')
    load_commonground(con, scrape_commonground_state_index.base)

    # DENORMALIZE

    # get all in one table
    title, data = get_all(con)
    con.close()

    #######################################


    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('data/clean/bridge.xlsx')
    worksheet = workbook.add_worksheet()


    # TODO Write some data headers.
    i = 0
    for j, val in enumerate(title):
        worksheet.write(i,j,val)


    for row in data:
        i += 1
        for j, val in enumerate(row):
            worksheet.write(i,j,val)

    workbook.close()


