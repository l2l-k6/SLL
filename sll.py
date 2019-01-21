#!/usr/bin/python
# -*- coding: utf-8 -*-

def get_contacts(players_db_f, opponents):
	'''Given the open XLS file players_db and the list of opponents,
	obtain the contact details (phone numbers and email address) of
	the opponenets. Sort the phone numbers by reachability, i.e.
	mobile < home < office < empty. Return a list of pairs (first
	name, last name, email, phone_0, phone_1, phone_2).
	'''

	# Squash Center Karlsruhe still uses the old format.
	import xlrd
	from mmap import mmap, ACCESS_READ

	with players_db_f as f:
		players_db = mmap(f.fileno(), 0, access = ACCESS_READ)
		players_wb = xlrd.open_workbook(file_contents =
			players_db)
		players_sh = players_wb.sheet_by_index(0)

		contacts = []
	        for o in opponents:
			name, prefix = parse_player_name(o)
			row, col = xls_search_player(players_sh,
				name, prefix)
			contacts.append(
				[players_sh.cell(row, col + 1).value,
				name,
				players_sh.cell(row, col + 5).value,
				players_sh.cell(row, col + 4).value,
				players_sh.cell(row, col + 2).value,
				players_sh.cell(row, col + 3).value]
			)

       	return contacts


def get_group_opponents(league_db_f, player):
	'''Given the open XLS file league_db_f and player’s name player,
	obtain the name of player’s group and the list of his opponents
	from the XLS file. Return the pair of these values.
	'''

	# Squash Center Karlsruhe still uses the old format.
	import xlrd
	from mmap import mmap, ACCESS_READ

	with league_db_f as f:
		league_db = mmap(f.fileno(), 0, access = ACCESS_READ)

		# Read table data:
		league_wb = xlrd.open_workbook(file_contents=league_db)
		league_sh = league_wb.sheet_by_index(0)

		# Get coordinates of the cell containing player’s name:
		matches = xls_search_str(league_sh, player)
		if len(matches) == 0:
			raise ValueError('Player "%s" not found in the ' +
				'database.' % player)
		if len(matches) >= 2:
			raise ValueError('More than one player "%s" '
				'found in the database.' % player)
		row, col = matches[0]

		# Get vertical bounds of the group of the player:
		row_low, row_high = xls_find_table_borders(league_sh, row, col)

		# Get the group’s name:
		group = league_sh.cell(row_low - 1, col - 1).value

		# Get the names of the opponents:
		opponents = []
		from itertools import chain
		for i in chain(range(row_low, row),
			range(row + 1, row_high + 1)):
				opponents.append(league_sh.cell(i, col).value)

	return (group, opponents)


def parse_player_name(player_name):
	'''Split the string player_name into a string consisting of
	player’s last name and a string consisting of a prefix of
	player’s first name, remove punctuation and return this pair
	(last_name, first_name_prefix).

	For example 'Meyer M.' => ['Meyer', 'M'],
	'Meyer Th' => ['Meyer', 'Th'] or 'Meyer' => ['Meyer', ''].
	'''

	split_string = player_name.split()

	if len(split_string) == 0:
		return ('', '')
        if len(split_string) == 1:
		return (split_string[0], '')
	if len(split_string) > 2:
		raise ValueError('"%s" is not a valid player name.' % player_name)

	last_name = split_string[0]
	first_name_prefix = split_string[1].replace('.', '')
	return (last_name, first_name_prefix)

def parse_cmd_line():
	''' Parse command line arguments to get the player name and
	files containing the league database and players database.
	Return an instante of argparse.Namespace containing these
	arguments.
	'''

	from argparse import ArgumentParser, FileType

	description = 'Compile a list of opponents and how to ' \
		'contact them.'
	parser = ArgumentParser(description=description)
	parser.add_argument('-n', '--name', default='Chaichenets',
		help='Player for whom to look up the opponents. '
			'Defaults to %(default)s.',
		metavar='NAME')
	parser.add_argument('league_db',
		help='File containing the database of the current ' + \
			'league in XLS (Microsoft Excel) file format.',
		metavar='LEAGUE_DB',
		type=FileType(mode='rb'))
	parser.add_argument('players_db',
		default='Rangliste Adressen.xls',
		help='File containing the database of players in ' + \
			'XLS file format. Defaults to "%(default)s".',
		metavar='PLAYERS_DB',
		nargs='?',
		type=FileType(mode='rb'))

	return parser.parse_args()

def xls_find_table_borders(xls_sheet, row, col):
	u'''Search in the XLS sheet xls_sheet for the boundary entries
	of the table containing the cell with coordinates (row, col).
	Return a pair of the lowest and highest vertical coordinate
	(low_row, high_row) of the table.

	The table is assumed to be bounded vertically by empty cells.
	'''

	from xlrd import XL_CELL_EMPTY

	low_row = row
	while(xls_sheet.cell(low_row, col).ctype != XL_CELL_EMPTY):
		low_row -= 1

	high_row = row
	while(xls_sheet.cell(high_row, col).ctype != XL_CELL_EMPTY):
		high_row += 1

	return (low_row + 1, high_row - 1)

def xls_search_player(xls_sheet, last_name, prefix):
	u'''Search in the XLS sheet xls_sheet for a player entry
	characterized by the last name last_name and the fact that the
	first name begins with prefix.

	Return the pair (row, col) of the match. Raise ValueError, if
        there are no matches or if the match is not unique.

	The column consisting of first names is assumed to be the column
	next to the right of the column of last names.
	'''

	last_name_matches = xls_search_str(xls_sheet, last_name)
	prefix_matches = []
	for (m_row, m_col) in last_name_matches:
		if xls_sheet.cell(m_row, m_col + 1).value.startswith(prefix):
			prefix_matches.append((m_row, m_col))
	if len(prefix_matches) == 0:
		raise ValueError('No matches found.')
	if len(prefix_matches) > 1:
		raise ValueError('More than one (%i) match found.' % len(prefix_matches))
	return prefix_matches[0]

def xls_search_str(xls_sheet, search_str):
	u'''Search in the XLS sheet xls_sheet for the string
	search_str and return a list of coordinate pairs
	[(row, col)] of all mathes.

	Redundant whitespace is ignored in the XLS sheet for the
	string comparison. The search string, however, is taken as is.
	'''

	from xlrd import XL_CELL_TEXT

	matches = []
	for i in range(xls_sheet.nrows):
		for j in range(xls_sheet.ncols):
			if xls_sheet.cell(i, j).ctype == XL_CELL_TEXT:
				cmp_val = ' '.join(xls_sheet.cell(i, j).value.split())
				if search_str == cmp_val:
					matches.append((i, j))

	return matches

if __name__ == '__main__':

	# Get player's name and files containing the league and players
	# databases.
	args = parse_cmd_line()

	# Get player’s group and a list of his opponents:
	group, opponents = get_group_opponents(args.league_db,
		args.name)

	# Get the opponents’ contacts:
	contacts = get_contacts(args.players_db, opponents)

	# Display obtained data:
	print('In this season you play in group %s.\n' % group)
	print('Your opponents are as follows:')
	print('Last name, first name \t Mobile \t Work \t Home \t E-Mail')
	for c in contacts:
		print('%s, %s \t %s \t %s \t %s \t %s' % c)

