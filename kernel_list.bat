pip install jupyterlab
pip install octave_kernel
jt -t oceans16 -f fira -fs 9 -nf ptsans -nfs 7 -N -kl -cursw 3 -cursc r -cellw 95% -T
pip install ipywidgets
jupyter nbextension enable --py widgetsnbextension 
pip install jupyter_contrib_nbextensions && jupyter contrib nbextension install

ipython kernelspec list
pause