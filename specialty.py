class Specialty:
    def __init__(self, name):
        self.name = name
        self.recruiter_folders = []
        self.recruiter_dictionary = {}
        self.subfolders = ['1opens', '2clicks', '3bounces']
        self.prefix = ''
        if self.name == "Neurology":
            self.recruiter_folders = ['marissa', 'samantha', 'andy', 'tori', 'jason', 'adam', 'jennifer', 'hannah', 'connor']
            self.recruiter_dictionary = {'marissa': 'Marissa Whalen', 'samantha': 'Samantha Dorazio', 'andy': 'Andy Fadenholz', 'tori': 'Tori Gould',
                                         'jason': 'Jason Hermanutz', 'adam': 'Adam Hine', 'jennifer': 'Jennifer Sereda',
                                         'hannah': 'Hannah Watene', 'connor': 'Connor Olczak'}
            self.prefix = 'NL'
        elif self.name == "Urology":
            self.recruiter_folders = ['chelsea', 'scott', 'sandy', 'rasheeda']
            self.recruiter_dictionary = {'chelsea': 'Chelsea Burgess', 'scott': 'Scott Greenberg', 'sandy': 'Sandy Oliveaux', 'rasheeda': 'Rasheeda Scott'}
            self.prefix = 'UL'
        elif self.name == 'Gastro':
            self.recruiter_folders = ['kevin']
            self.recruiter_dictionary = {'kevin': 'Kevin Morgan'}
            self.prefix = 'GI'
        elif self.name == "Neurosurgery":
            self.subfolders.append('4unsubscribed')
            self.recruiter_folders = ['nancy', 'jonathan', 'rachel', 'andrea']
            self.recruiter_dictionary = {'andrea': 'Andrea Winslow', 'jonathan': 'Jonathan Haines', 'nancy': 'Nancy Cusick',
                              'rachel': 'Rachel Prero'}
            self.prefix = 'NS'


