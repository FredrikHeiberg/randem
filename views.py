from flask import flash, redirect, render_template, request, session, url_for
from routes import app
from flask.ext.bcrypt import Bcrypt
bcrypt = Bcrypt(app)
from functools import wraps


@app.route('/login', methods=['GET', 'POST'])
def login():
	error = None

	form = LoginForm(request.form)

	if request.method == 'POST':
		if form.validate_on_submit():
			user = User.query.filter_by(name=request.form['username']).first()
			if user is not None and bcrypt.check_password_hash(
				user.password, request.form['password']):
				session['logged_in'] = True
				flash('Du er naa logget inn!')
				return redirect(url_for('index'))
			else:
				error = 'Feil brukernavn/passord'
	return render_template('login.html', form=form, error=error)

@app.route('/logout')
@login_required
def logout():
	session.pop('logged_in', None)
	flash('Du er naa logget ut!')
	return redirect(url_for('index'))