import json

import logging

import plac
import random
import warnings
from pathlib import Path
import spacy
from spacy.util import minibatch, compounding


class SpacyModel:

    def __init__(self):
        logging.info('Called SpacyModel')

    def jsonToTupple(self, TrainingDataPath):
        """
        ----------
        Function
        ----------
        * Takes a JSON file for training SpCy model and returns a SpaCy readable tuple

        --------
        INPUT
        --------
        TrainingDataPath: (str) location to the training JSON File

        -------
        RETURN
        -------
        trainingData: Tuple that will be used to train spacy model
        """

        logging.info('Entering SpacyModel.jsonToTupple')

        with open(TrainingDataPath) as train_data:
            train = json.load(train_data)
        trainingData = []
        for data in train:
            ents = [tuple(entity) for entity in data['entities']]
            trainingData.append((data['content'], {'entities': ents}))
        with open('{}'.format(TrainingDataPath.replace('json', 'txt')), 'w') as write:
            write.write(str(trainingData))

        logging.info('Exiting SpacyModel.jsonToTupple')
        return trainingData

    def labelTraining(self, trainingData, n_iter,
                      output_dir='TrainedModel',
                      modelName='CustomSpacyModel',
                      model=None):
        """
        ----------
        Function
        ----------
        * Train the NLP label detection model based on given data
        * Return a trained model

        --------
        INPUT
        --------
        trainingData: (Tuple): Text and labeled entity
        n_iter: (int): Number of iterations to perform for training the model
        output_dir: (str): TrainedModel (Default)
                                Location to save trained model
        modelName: (str): CustomSpacyModel (Default)
                            Name of the newly trained model
        model: (Spacy Model): spacy.blank('en') - Blank by default
                            Takes input of that spay model that want to be trained

        -------
        RETURN
        -------
        nlp: SpaCy nlp model for labeling
        """

        logging.info('Entering SpacyModel.LabelTraining')

        TRAIN_DATA = trainingData

        if model is not None:
            nlp = spacy.load(model)  # load existing spaCy model
            print("Loaded model '%s'" % model)
        else:
            nlp = spacy.blank("en")  # create blank Language class
            print("Created blank 'en' model")

        # create the built-in pipeline components and add them to the pipeline
        # nlp.create_pipe works for built-ins that are registered with spaCy
        if "ner" not in nlp.pipe_names:
            ner = nlp.create_pipe("ner")
            nlp.add_pipe(ner, last=True)
        # otherwise, get it so we can add labels
        else:
            ner = nlp.get_pipe("ner")

        # add labels
        for _, annotations in TRAIN_DATA:
            for ent in annotations.get("entities"):
                ner.add_label(ent[2])

            # get names of other pipes to disable them during training
            pipe_exceptions = ["ner", "trf_wordpiecer", "trf_tok2vec"]
            other_pipes = [pipe for pipe in nlp.pipe_names if pipe not in pipe_exceptions]
            # only train NER
            with nlp.disable_pipes(*other_pipes), warnings.catch_warnings():
                # show warnings for misaligned entity spans once
                warnings.filterwarnings("once", category=UserWarning, module='spacy')

                # reset and initialize the weights randomly – but only if we're
                # training a new model
                if model is None:
                    nlp.begin_training()
                for itn in range(n_iter):
                    random.shuffle(TRAIN_DATA)
                    losses = {}
                    # batch up the examples using spaCy's minibatch
                    batches = minibatch(TRAIN_DATA, size=compounding(4.0, 32.0, 1.001))
                    for batch in batches:
                        texts, annotations = zip(*batch)
                        nlp.update(
                            texts,  # batch of texts
                            annotations,  # batch of annotations
                            drop=0.5,  # dropout - make it harder to memorise data
                            losses=losses,
                        )
                    print("Losses", losses)

            # test the trained model
            for text, _ in TRAIN_DATA:
                doc = nlp(text)
                print("Entities", [(ent.text, ent.label_) for ent in doc.ents])
                print("Tokens", [(t.text, t.ent_type_, t.ent_iob) for t in doc])

            # save model to output directory
            if output_dir is not None:
                output_dir = Path(output_dir)
                if not output_dir.exists():
                    output_dir.mkdir()
                nlp.to_disk(output_dir)
                print("Saved model to", output_dir)

        logging.info('Exiting SpacyModel.labelTraining')
        return nlp

    def useModel(self, pathToTrainedModel):
        """
        ----------
        Function
        ----------
        * Loads the trained model

        --------
        INPUT
        --------
        pathToTrainedModel: (str) location to the trained model

        -------
        RETURN
        -------
        trainingData: Tuple that will be used to train spacy model
        """

        logging.info('Entering SpacyModel.useModel')

        nlp = spacy.load(pathToTrainedModel)
        return nlp

        logging.info('Exiting SpacyModel.useModel')

    def labelMaker(self,
                   textData,
                   SpacyModel = spacy.load('en_core_web_sm')):
        """
        ----------
        Function
        ----------
        * Used for making label prediction of the text data.

        --------
        INPUT
        --------
        SpacyModel: (Spacy trained model
        textData: Data who's label need to be predicted

        -------
        RETURN
        -------
        modelResultDict: (Dict) returns a dict of dict type object with label as the textData and
                                key as  a dict of start and end location and predicted label of the text.
        """

        modelResultDict = {}
        Modelresult = SpacyModel(textData)
        for ent in Modelresult.ents:
            modelResultDict[ent.text] = {"start_char": ent.start_char,
                                         'end_char': ent.end_char,
                                         'label': ent.label_}
        return modelResultDict
