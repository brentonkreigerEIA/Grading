# add this to each pass through, that will be easier
import pandas as pd
import numpy as np

a = np.array([[0, 'Duke', 2], [3, 'USI', 5], [5, '', 8], [9, 'Boulder', 1]])
df1 = pd.DataFrame(a, columns=['First', 'University Name', 'Second'])
ind = np.array(range(0, len(df1)))

# for k in ind:
#     if len(df1.iloc[k, 1]) > 2:
#         print(df1.iloc[k, 1])


print(df1.iloc[1, :])

#use notnull