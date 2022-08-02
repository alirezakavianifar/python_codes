import pandas as pd
import seaborn as sns
import numpy as np
import matplotlib.pyplot as plt

df = pd.read_csv(r'E:\Tutorials\Data\opt-digits\optdigits_csv.csv')

df.columns = ["P" + str(i) for i in range(0, len(df.columns) - 1)] + ["y"]

df = df.loc[df.y.isin([1,3,6])]

from sklearn.model_selection import train_test_split

X_train, X_test, y_train, y_test = train_test_split(
    df.filter(regex = '\d'),
    df.y,
    test_size=0.30,
    random_state=1)

trn = X_train
trn['y'] = y_train

tst = X_test
tst['y'] = y_test

fig, ax = plt.subplots(
    nrows=1,
    ncols=20,
    figsize=(15, 3.5),
    subplot_kw=dict(xticks=[], yticks=[]))
    
for i in np.arange(20):
    ax[i].imshow(X_train.to_numpy()[i, 0:-1].reshape(8,8), cmap=plt.cm.gray)
plt.show()    

g = sns.PairGrid(
    trn,
    vars=["P25","P30","P45","P60"],
    hue="y",
    diag_sharey=(False),
    palette=["red","green","blue"])

g.map_diag(plt.hist)

g.map_upper(sns.kdeplot)

g.map_lower(sns.scatterplot)

g.add_legend()

df.to_csv(r'E:\Tutorials\Data\opt-digits/optidigits.csv', sep=',', index=False)
trn.to_csv(r'E:\Tutorials\Data\opt-digits/optidigits_trn.csv', sep=',', index=False)
tst.to_csv(r'E:\Tutorials\Data\opt-digits/optidigits_tst.csv', sep=',', index=False)

from sklearn.decomposition import PCA

pca = PCA()
trn_ft = pca.fit_transform(trn)
plt.plot(pca.explained_variance_ratio_)

tst_tf = pca.transform(X_test)         
sns.scatterplot(
    x=trn_ft[:, 0],
    y=trn_ft[:, 1],
    style=y_train,
    hue=y_train,
    palette=['red','green','blue'])
    






   

