import argparse
from word_common import load_profile

p = argparse.ArgumentParser(description='Validate a formatting profile without starting Word.')
p.add_argument('profile')
a = p.parse_args()
profile = load_profile(a.profile)
print('Valid schema version 1 profile:', profile.get('name', a.profile))
