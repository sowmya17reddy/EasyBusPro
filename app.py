from flask import Flask, render_template, request, redirect, session
import sqlite3

app = Flask(__name__)
app.secret_key = "easybuspro123"

@app.route('/')
def home():
    return render_template('landing.html')

@app.route('/signin', methods=['GET', 'POST'])
def signin():

    if request.method == 'POST':

        email = request.form['email']
        password = request.form['password']

        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()

        cursor.execute("""
            SELECT * FROM users
            WHERE email=? AND password=?
        """, (email, password))

        user = cursor.fetchone()

        conn.close()

        if user:
            session["user_id"] = user[0]
            session["user_name"] = user[1]
            session["user_email"] = user[2]
            return redirect('/dashboard')
        else:
            return "Invalid Email or Password!"

    return render_template("signin.html")

@app.route('/signup', methods=['GET', 'POST'])
def signup():

    if request.method == 'POST':

        name = request.form['name']
        email = request.form['email']
        phone = request.form['phone']
        password = request.form['password']
        confirm_password = request.form['confirm_password']

        if password != confirm_password:
            return "Passwords do not match!"

        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()

        try:
            cursor.execute("""
                INSERT INTO users(name,email,phone,password)
                VALUES(?,?,?,?)
            """, (name, email, phone, password))

            conn.commit()

        except sqlite3.IntegrityError:
            conn.close()
            return "Email already registered!"

        conn.close()

        return redirect('/signin')

    return render_template("signup.html")

@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

@app.route('/results')
def results():

    from_city = request.args.get('from')
    to_city = request.args.get('to')
    date = request.args.get('date')
    passengers = request.args.get('passengers')

    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT * FROM bus
        WHERE from_city=? AND to_city=?
    """, (from_city, to_city))

    buses = cursor.fetchall()

    conn.close()

    return render_template(
        "results.html",
        buses=buses,
        from_city=from_city,
        to_city=to_city,
        date=date,
        passengers=passengers
    )

@app.route('/search')
def search():
    return render_template('search.html')

@app.route('/seat_selection')
def seat():
    bus_name = request.args.get('bus')
    from_city = request.args.get('from')
    to_city = request.args.get('to')
    date = request.args.get('date')
    passengers = request.args.get('passengers')
    price = request.args.get('price')

    return render_template(
        'seat_selection.html',
        bus_name=bus_name,
        from_city=from_city,
        to_city=to_city,
        date=date,
        passengers=passengers,
        price=price
    )

@app.route('/payment')
def payment():
    bus_name = request.args.get('bus')
    from_city = request.args.get('from')
    to_city = request.args.get('to')
    date = request.args.get('date')
    passengers = request.args.get('passengers')
    seats = request.args.get('seats')
    total = request.args.get('total')

    return render_template(
        'payment.html',
        bus_name=bus_name,
        from_city=from_city,
        to_city=to_city,
        date=date,
        passengers=passengers,
        seats=seats,
        total=total
    )

@app.route('/success')
def success():

    # Check if the user is logged in
    if "user_id" not in session:
        return redirect("/signin")

    bus_name = request.args.get('bus')
    from_city = request.args.get('from')
    to_city = request.args.get('to')
    date = request.args.get('date')
    passengers = request.args.get('passengers')
    seats = request.args.get('seats')
    total = request.args.get('total')
    payment_method = request.args.get('payment_method')

    conn = sqlite3.connect("database.db")
    cursor = conn.cursor()

    # Get Bus ID
    cursor.execute(
        "SELECT id FROM bus WHERE bus_name=?",
        (bus_name,)
    )

    bus = cursor.fetchone()

    if bus:

        bus_id = bus[0]

        cursor.execute("""
            INSERT INTO bookings
            (user_id, bus_id, journey_date,
             seat_numbers, passengers,
             amount, payment_method)

            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            session["user_id"],
            bus_id,
            date,
            seats,
            passengers,
            total,
            payment_method
        ))

        conn.commit()

    conn.close()

    return render_template(
        "success.html",
        bus_name=bus_name,
        from_city=from_city,
        to_city=to_city,
        date=date,
        passengers=passengers,
        seats=seats,
        total=total,
        payment_method=payment_method
    )

@app.route('/my_bookings')
def my_bookings():

    # Check whether the user is logged in
    if "user_id" not in session:
        return redirect('/signin')

    # Connect to the database
    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # Fetch all bookings of the logged-in user
    cursor.execute("""
        SELECT
            bookings.booking_id,
            bus.bus_name,
            bus.from_city,
            bus.to_city,
            bookings.journey_date,
            bookings.seat_numbers,
            bookings.passengers,
            bookings.amount,
            bookings.payment_method

        FROM bookings

        JOIN bus
        ON bookings.bus_id = bus.id

        WHERE bookings.user_id = ?

        ORDER BY bookings.booking_time DESC
    """, (session["user_id"],))

    bookings = cursor.fetchall()

    conn.close()

    return render_template(
        "my_bookings.html",
        bookings=bookings
    )

@app.route('/profile')
def profile():

    if "user_id" not in session:
        return redirect('/signin')

    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute(
        "SELECT * FROM users WHERE id=?",
        (session["user_id"],)
    )

    user = cursor.fetchone()

    conn.close()

    return render_template(
        "profile.html",
        user=user
    )

@app.route('/admin', methods=['GET', 'POST'])
def admin():

    if request.method == 'POST':

        bus_name = request.form['bus_name']
        from_city = request.form['from_city']
        to_city = request.form['to_city']
        departure = request.form['departure']
        arrival = request.form['arrival']
        price = request.form['price']
        available_seats = request.form['available_seats']

        conn = sqlite3.connect("database.db")
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO bus
            (bus_name, from_city, to_city,
             departure, arrival, price, available_seats)

            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            bus_name,
            from_city,
            to_city,
            departure,
            arrival,
            price,
            available_seats
        ))

        conn.commit()
        conn.close()

        return redirect('/admin')

    return render_template("admin.html")

@app.route('/admin_dashboard')
def admin_dashboard():

    conn = sqlite3.connect("database.db")
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM bus")

    buses = cursor.fetchall()

    conn.close()

    return render_template(
        "admin_dashboard.html",
        buses=buses
    )

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)