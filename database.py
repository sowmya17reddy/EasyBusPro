import sqlite3

# Connect to SQLite database
conn = sqlite3.connect("database.db")

# Create a cursor
cursor = conn.cursor()

# -------------------------
# Create Users Table
# -------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    phone TEXT,
    password TEXT NOT NULL
)
""")

# -------------------------
# Create Bus Table
# -------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS bus (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bus_name TEXT NOT NULL,
    from_city TEXT NOT NULL,
    to_city TEXT NOT NULL,
    departure TEXT,
    arrival TEXT,
    price INTEGER,
    available_seats INTEGER
)
""")

# -------------------------
# Create Bookings Table
# -------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS bookings (
    booking_id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    bus_id INTEGER,
    journey_date TEXT,
    seat_numbers TEXT,
    passengers INTEGER,
    amount INTEGER,
    payment_method TEXT,
    booking_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
""")

cursor.execute("""
INSERT OR IGNORE INTO bus
(bus_name, from_city, to_city, departure, arrival, price, available_seats)
VALUES
('Royal Travels', 'Hyderabad', 'Bangalore', '08:00 PM', '06:00 AM', 1050, 32)
""")

cursor.execute("""
INSERT OR IGNORE INTO bus
(bus_name, from_city, to_city, departure, arrival, price, available_seats)
VALUES
('Orange Travels', 'Hyderabad', 'Bangalore', '09:30 PM', '07:15 AM', 1200, 25)
""")

cursor.execute("""
INSERT OR IGNORE INTO bus
(bus_name, from_city, to_city, departure, arrival, price, available_seats)
VALUES
('VRL Travels', 'Hyderabad', 'Bangalore', '10:00 PM', '07:30 AM', 980, 18)
""")

# Save changes
conn.commit()

# Close connection
conn.close()

print("Database and tables created successfully!")